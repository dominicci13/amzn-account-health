"""Weekday Amazon Account Health scrape.

Once per weekday (Mon-Fri 11:00 local) this script:

1. Opens `AH-Metrics.xlsm` hidden and clears any pre-existing chart shapes
   on the dashboard sheet via the `modUtilities.deleteCharts` macro.
2. For each Amazon account in `AMAZON_URLS`, logs into Seller Central and
   scrapes the Account Health (Late Shipment Rate, Pre-Fulfillment Cancel
   Rate, Valid Tracking Rate) + Premium Shipping Options (eligibility, OTDR,
   Cancellation Rate, Valid Tracking Rate) widgets.
3. For each size tier (Standard / Oversize), scrapes the Seller Fulfilled
   Prime page — program status, the speed-distribution chart (1d / 2d / 2d+),
   and the Fulfillment + Supporting metrics tables — pasting the chart
   screenshot into the dashboard via the Windows clipboard.
4. Calls `modUtilities.resizeCharts` to normalize chart dimensions across
   all pasted images, stamps today's date into both sheets' header cells,
   and saves + closes the workbook.
5. Emails the refreshed workbook as an `.xlsm` attachment via Outlook.

Both chart macros return "" or a failure reason (never a `MsgBox`, which would hang
the hidden Excel) and run under a watchdog: a reason, or a macro still running after
`MACRO_TIMEOUT_SEC`, crashes the run into `handle_crash` instead of emailing a
half-built dashboard. The Excel this run opens is ended by its own pid on every path,
success or crash; nothing is ever killed by name.
"""
from __future__ import annotations

import io
import os
import signal
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import pywintypes
import win32api
import win32clipboard
import win32con
import xlwings as xw
from dotenv import load_dotenv
from PIL import Image
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from seller_automation_utils import WorkbookRefreshError, accounts, alert_utils, chrome, custom_functions, greeting_for, outlook
from seller_automation_utils.config_utils import get_env, load_config_safe
from seller_automation_utils.logging_utils import setup_logging
from seller_automation_utils.schedule_utils import run_on_schedule
from seller_automation_utils.ui_utils import ask_user


log = setup_logging("amzn_account_health")
load_dotenv()
username: str = get_env("AMZN_email", required=True)
password: str = get_env("AMZN_pass", required=True)
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)
sender_email: str = get_env("SENDER_EMAIL", required=True)
to_email: list[str] = [e.strip() for e in (get_env("TO_EMAIL", required=True) or "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in (get_env("CC_EMAIL", default="") or "").split(",") if e.strip()]

_accounts_cfg = load_config_safe(Path(__file__).resolve().parent / "config" / "accounts.json")
_metrics_columns: dict[str, str] = _accounts_cfg.get("account_health_metrics_columns", {})
_dashboard: dict[str, dict] = _accounts_cfg.get("account_health_dashboard", {})

_layout = load_config_safe(Path(__file__).resolve().parent / "config" / "metrics_layout.json")
_metrics_rows: dict[str, int] = _layout.get("metrics_rows", {})

_paths = load_config_safe(Path(__file__).resolve().parent / "config" / "paths.json")
ah_wb_path: str = _paths["ah_wb_path"]


# The 2026-07-30 Program Eligibilities redesign moved each program onto its own
# widget page, iframed into /ah/eligibility. Loading the widget directly avoids
# having to switch frames, and is what the old /performance/eligibilities URL now
# redirects into.
PSO_WIDGET_URL = "https://sellercentral.amazon.com/fbm-seller-program-dashboard/widget?programId=PSO"

# That redesign also dropped the "2411/2413" order counts that used to sit beside
# each rate. Nothing on the page carries them any more, and the downloadable defect
# report lists only defective orders, so the denominator is unrecoverable.
COUNT_UNAVAILABLE = "N/A"

# Opening the workbook plus deleteCharts took at most 23s over 97 logged runs, and
# resizeCharts plus the save at most 16s; 3x that is under the fleet's 300s floor.
MACRO_TIMEOUT_SEC: int = 300


def _premium_shipping_values(status: str, metrics: dict[str, str]) -> list[str]:
    """Order Premium Shipping widget readings to match metrics sheet rows 17-23.

    The widget lists its metrics as On-Time Delivery, Valid Tracking, Cancellation,
    which is not the workbook's row order, so values are looked up by title rather
    than by position.

    Args:
        status: Eligibility badge label, e.g. `"Eligible"`.
        metrics: Metric title -> current rate, as read off the widget card.

    Returns:
        Seven values for rows 17-23: eligibility status, on-time delivery rate,
        its count cell, cancellation rate, its count cell, valid tracking rate,
        its count cell. Count cells are always `COUNT_UNAVAILABLE`; a metric the
        widget did not render falls back to it too.
    """
    return [
        status or COUNT_UNAVAILABLE,
        metrics.get("On-Time Delivery Rate", COUNT_UNAVAILABLE),
        COUNT_UNAVAILABLE,
        metrics.get("Cancellation Rate", COUNT_UNAVAILABLE),
        COUNT_UNAVAILABLE,
        metrics.get("Valid Tracking Rate", COUNT_UNAVAILABLE),
        COUNT_UNAVAILABLE,
    ]


def _macro_failure(result: object) -> str | None:
    """Read a chart macro's return value as success (``None``) or a failure reason.

    ``""`` is the hardened ``Function``'s success. ``None`` is what an old ``Sub`` in the
    workbook returns, so it is success too, which keeps this safe before the workbook
    module is re-pasted. Anything else is not a value the contract defines, so it is a
    failure rather than a guess.

    Args:
        result: What ``wb.macro(name)()`` returned.

    Returns:
        ``None`` on success, otherwise the reason to raise with.
    """
    if result is None or result == "":
        return None
    if isinstance(result, str):
        return result
    return f"returned an unexpected {type(result).__name__} ({result!r}); expected '' or a failure reason"


def _pin_pid(pid: int) -> object | None:
    """Hold a handle on this run's Excel so Windows cannot give its pid to another process.

    ``_kill_pid`` and ``excel.kill()`` both act on a bare pid. While this handle is open,
    that pid keeps naming this Excel even after it has exited, so a kill that lands late
    fails on the dead process ("Access is denied") instead of reaching whatever program
    would otherwise have been handed the number. Close it only after the last such kill.

    Args:
        pid: Process id of the Excel instance this run has just started.

    Returns:
        The open handle, or ``None`` if it could not be opened.
    """
    try:
        return win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except pywintypes.error as exc:
        log.warning(f"Could not hold Excel pid {pid} ({exc}); a late kill is no longer pid-reuse safe.")
        return None


def _kill_pid(pid: int) -> None:
    """Terminate one process by pid, ignoring one that has already gone.

    ``os.kill`` on Windows is ``TerminateProcess`` on exactly this pid. It is used from the
    watchdog thread because the xlwings App cannot be touched there: any property on it
    is a COM call into the very Excel that is stuck. ``main`` holds the pid with
    ``_pin_pid`` for as long as this can fire, so it cannot hit a reused pid.

    Args:
        pid: Process id of the Excel instance this run started.
    """
    try:
        os.kill(pid, signal.SIGTERM)
    except OSError as exc:
        log.info(f"Excel pid {pid} was already gone ({exc}).")


def _run_checked_macro(excel: xw.App, wb: xw.Book, macro_name: str) -> None:
    """Run a chart macro, raise on its failure reason, and bound it in time.

    A macro stuck on a modal dialog blocks this COM call for good, and no COM call can
    interrupt it. A watchdog timer therefore kills this run's Excel, by the pid read
    before the call, once ``MACRO_TIMEOUT_SEC`` passes; the blocked call then fails and
    is reported as the timeout it was.

    ``display_alerts`` is turned back off afterwards because the macro's Cleanup turns it
    on, and with alerts on a failed save raises a modal Save As instead of an error.

    Args:
        excel: The Excel instance this run started.
        wb: The open workbook holding the macro.
        macro_name: Full macro name, e.g. ``"modUtilities.deleteCharts"``.

    Raises:
        WorkbookRefreshError: The macro returned a failure reason, or overran
            ``MACRO_TIMEOUT_SEC`` and its Excel was killed.
        pywintypes.com_error: The macro call itself failed, within the time bound.
    """
    pid = excel.pid
    timeout = MACRO_TIMEOUT_SEC
    timed_out = threading.Event()

    def _on_timeout() -> None:
        timed_out.set()
        log.error(
            f"{macro_name} exceeded {timeout}s. Excel is stuck, most likely on a modal dialog. "
            f"Killing its pid {pid}."
        )
        _kill_pid(pid)

    watchdog = threading.Timer(timeout, _on_timeout)
    watchdog.daemon = True
    log.info(f"Running macro: [cyan]{macro_name}[/cyan]")
    watchdog.start()
    try:
        result = wb.macro(macro_name)()
    except pywintypes.com_error as exc:
        if timed_out.is_set():
            raise WorkbookRefreshError(f"{macro_name} exceeded {timeout}s, Excel was killed") from exc
        raise
    finally:
        watchdog.cancel()
        # cancel() does not stop a timeout that has already started; wait it out so
        # timed_out is final before it is read
        watchdog.join()

    if timed_out.is_set():
        # Returned just as the watchdog fired: Excel is gone either way
        raise WorkbookRefreshError(f"{macro_name} exceeded {timeout}s, Excel was killed")

    excel.display_alerts = False
    excel.screen_updating = False

    failure = _macro_failure(result)
    if failure is not None:
        log.error(f"{macro_name} reported: {failure}")
        raise WorkbookRefreshError(f"{macro_name}: {failure}")


def _end_excel(excel: xw.App) -> None:
    """Quit the Excel this run started, then make sure that process is gone.

    ``kill()`` targets the pid xlwings recorded when it launched this instance, so no
    other Excel on the machine is touched. It runs even when ``quit()`` fails, which is
    the case where the process would otherwise be left behind holding the workbook.
    After a clean quit it fails harmlessly on the exited process: the caller holds the
    pid with ``_pin_pid``, so it cannot have been reused. Since library 1.8.8
    ``handle_crash`` only ends Excel started through ``excel_utils``, so this is the only
    thing that ends this one after a crash.

    Args:
        excel: The ``xw.App`` this run created.
    """
    try:
        # With alerts on, quitting over an unsaved workbook (the crash path) raises a Save
        # prompt on a hidden Excel and hangs.
        excel.display_alerts = False
        excel.quit()
    except Exception as exc:
        log.warning(f"Excel did not quit cleanly ({exc}). Killing it.")
    try:
        excel.kill()
    except Exception:
        pass  # already exited after quit()


def _email_body() -> str:
    """Build the email body with the time-of-day greeting."""
    return (
        f"{greeting_for()},<br><br>"
        "Please find attached Account Health Metrics file updated for today.<br><br>"
        "If any questions, please let me know.<br><br>"
        "Thanks,<br><br>"
    )


def main() -> None:
    """Scrape account health metrics from each Amazon account and email the workbook.

    Opens the AH-Metrics workbook, loops through each Amazon account to scrape
    shipping performance, Prime eligibility, and Seller Fulfilled Prime stats,
    pastes charts into the dashboard sheet, saves the workbook, and emails it.
    """
    driver = None
    excel = None
    excel_pin = None
    try:
        log.info("Opening workbook and removing charts.")
        excel = xw.App(visible=False)
        excel_pin = _pin_pid(excel.pid)
        excel.display_alerts = False
        excel.screen_updating = False
        ah_wb = excel.books.open(ah_wb_path)
        sh_metrics = ah_wb.sheets(1)
        sh_dash = ah_wb.sheets(2)
        _run_checked_macro(excel, ah_wb, "modUtilities.deleteCharts")

        driver = chrome.start_browser(user_data_dir, "Default", headless=True)

        for account, root, url in accounts.iter_amazon_accounts():
            col = _metrics_columns[account]
            dash = _dashboard[account]

            log.info(f"Navigating to [cyan]{root}[/cyan] account.")
            driver.get(url)
            driver.switch_to_window(0)

            try:
                accounts.amazon_login(driver, username, username, password, retry_url=url)
            except TimeoutException:
                pass

            # Account health stats
            trying = True
            while trying:
                try:
                    driver.get("https://sellercentral.amazon.com/performance/dashboard")
                    driver.switch_to_window(0)

                    log.info(f"Getting [cyan]{root}[/cyan] Account Health Statistics.")
                    ship_rate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-late-shipment-rate-row"))).text.split("\n")
                    pre_cancel_rate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-cancellation-rate-row"))).text.split("\n")
                    valid_track_rate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-view-tracking-rate-row"))).text.split("\n")
                    trying = False
                except TimeoutException:
                    log.error("Failed to get Account Health Statistics. Retrying.")
                    time.sleep(5)

            try:
                lsr: str = ship_rate[2]
                lsr_orders: str = ship_rate[3].replace("orders", "orders (30 days)")
            except IndexError:
                lsr, lsr_orders = "N/A", "N/A"

            try:
                pcr: str = pre_cancel_rate[2]
                pcr_orders: str = pre_cancel_rate[3].replace("orders", "orders (7 days)")
            except IndexError:
                pcr, pcr_orders = "N/A", "N/A"

            try:
                vtr_rate: str = valid_track_rate[2]
                vtr_orders: str = valid_track_rate[3].replace("orders", "orders (30 days)")
            except IndexError:
                vtr_rate, vtr_orders = "N/A", "N/A"

            values: list[str] = [lsr, lsr_orders, pcr, pcr_orders, vtr_rate, vtr_orders]
            df = pd.DataFrame([[v] for v in values])
            sh_metrics.range(f"{col}{_metrics_rows['account_health']}").value = df.values

            # Premium shipping stats
            driver.get(PSO_WIDGET_URL)
            driver.switch_to_window(0)

            log.info(f"Getting [cyan]{root}[/cyan] Premium Shipping Statistics.")
            card = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="program-card--PSO"]')))
            metric_css = '[data-testid="program-card__metrics__item"]'
            WebDriverWait(driver, 30).until(lambda d: card.find_elements(By.CSS_SELECTOR, metric_css))

            try:
                # kat-badge renders its text into a shadow root, so the label attribute is the only readable copy
                acc_status: str = card.find_element(By.CSS_SELECTOR, '[data-testid="program-card__status-badge"]').get_attribute("label") or ""
            except NoSuchElementException:
                acc_status = "Not Eligible"

            metrics: dict[str, str] = {}
            for item in card.find_elements(By.CSS_SELECTOR, metric_css):
                title: str = item.find_element(By.CSS_SELECTOR, '[data-testid="program-card__metrics__title"]').text.strip()
                # first line is a Katal status glyph, the rate is the last
                metrics[title] = item.find_element(By.CSS_SELECTOR, '[data-testid="program-card__metrics__current"]').text.split("\n")[-1].strip()

            values = _premium_shipping_values(acc_status, metrics)
            df = pd.DataFrame([[v] for v in values])
            sh_metrics.range(f"{col}{_metrics_rows['prime_performance']}").value = df.values

            # Seller Fulfilled Prime stats
            sfp_performance_urls = {
                "Standard-size": "https://sellercentral.amazon.com/seller-fulfilled-prime/seller-performance?tierName=STD",
                "Oversize": "https://sellercentral.amazon.com/seller-fulfilled-prime/seller-performance?tierName=OS"
            }

            for size, sfp_url in sfp_performance_urls.items():
                driver.get(sfp_url)
                driver.switch_to_window(0)
                time.sleep(3)

                size_key = "standard" if size == "Standard-size" else "oversize"

                log.info(f"Checking [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Program current status.")
                program_status: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "a-alert-content"))).text
                time.sleep(2)

                if program_status == "In Trial":
                    trial_end_date: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sfp-trial-performance-end-date"]'))).text.split(", ")
                    program_status = f"{program_status} (until {trial_end_date[0]})"
                elif program_status.startswith("Your Seller Fulfilled Prime performance has failed "):
                    program_status = "Revoked"

                sh_metrics.range(f"{col}{_metrics_rows[f'sfp_{size_key}_status']}").value = program_status

                # Speed metrics
                log.info(f"Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Speed metric charts.")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "show-gvs-metrics-by-date-filter-past-seven-days"))).click()
                element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "a-list-item")))
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(3)

                chart_css = f"#{'std' if size == 'Standard-size' else 'os'}-domestic-delivery-speed-graph > div > div:nth-child(1)"
                graph = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, chart_css)))

                # Screenshot the element itself rather than cropping a full-page capture.
                # The old crop fed element.location (document coordinates) into a viewport
                # screenshot, so the region slid by whatever the scroll offset happened to
                # be and some charts came out framing the legend instead of the plot.
                chart_img = Image.open(io.BytesIO(graph.screenshot_as_png))

                output = io.BytesIO()
                chart_img.convert("RGB").save(output, "BMP")
                data = output.getvalue()[14:]
                output.close()
                custom_functions.send_to_clipboard(win32clipboard.CF_DIB, data)

                try:
                    custom_functions.paste_image_from_clipboard(sh_dash, dash[f"{size_key}_chart_cell"])
                except Exception:
                    log.error("Failed to paste chart into Excel.")

                table_id = f"{'std' if size == 'Standard-size' else 'os'}-domestic-delivery-speed-details-table"
                log.info(f"Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Speed metrics Statistics.")
                raw_metrics: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, table_id))).text.split("\n")

                one_day: list[str] = raw_metrics[-3].split(" ")
                perc_one_day: str = one_day[3]
                views_one_day: str = one_day[-1]

                two_days: list[str] = raw_metrics[-2].split(" ")
                perc_two_days: str = two_days[3]
                views_two_days: str = two_days[-1]

                two_days_plus: list[str] = raw_metrics[-1].split(" ")
                perc_two_days_plus: str = two_days_plus[3]
                views_two_days_plus: str = two_days_plus[-1]

                speed_col = dash[f"{size_key}_speed_col"]
                speed_row = dash[f"{size_key}_speed_row"]
                sh_dash.range(f"{speed_col}{speed_row}").value = [perc_one_day, views_one_day]
                sh_dash.range(f"{speed_col}{speed_row + 2}").value = [perc_two_days, views_two_days]
                sh_dash.range(f"{speed_col}{speed_row + 4}").value = [perc_two_days_plus, views_two_days_plus]

                # Fulfillment metrics
                log.info(f"Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Fulfillment Statistics.")
                fulfillment_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "show-fulfillment-metrics-by-date-filter-past-seven-days")))
                driver.execute_script("arguments[0].scrollIntoView(true);", fulfillment_btn)
                fulfillment_btn.click()
                time.sleep(3)

                raw_otd: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "on-time-delivery.metrics-table-columns"))).text.split("\n")[1].split(" ")
                otd: str = raw_otd[0]
                otd_units: str = raw_otd[4]

                otd_row = _metrics_rows[f"sfp_{size_key}_otd"]
                sh_metrics.range(f"{col}{otd_row}").value = otd
                sh_metrics.range(f"{col}{otd_row + 1}").value = otd_units

                raw_vtr: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "valid-tracking.metrics-table-columns"))).text.split("\n")
                sfp_vtr: str = raw_vtr[1]
                sfp_vtr_packages: str = raw_vtr[3].split(" ")[0]

                vtr_row = _metrics_rows[f"sfp_{size_key}_vtr"]
                sh_metrics.range(f"{col}{vtr_row}").value = sfp_vtr
                sh_metrics.range(f"{col}{vtr_row + 1}").value = sfp_vtr_packages

                raw_cr: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "cancellation.metrics-table-columns"))).text.split("\n")
                cr: str = raw_cr[1]
                cr_units: str = raw_cr[3].split(" ")[0]

                cr_row = _metrics_rows[f"sfp_{size_key}_cr"]
                sh_metrics.range(f"{col}{cr_row}").value = cr
                sh_metrics.range(f"{col}{cr_row + 1}").value = cr_units

                # Supporting metrics
                log.info(f"Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Supporting Statistics.")
                support_btn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "show-supporting-metrics-by-date-filter-past-seven-days")))
                driver.execute_script("arguments[0].scrollIntoView(true);", support_btn)
                support_btn.click()
                time.sleep(3)

                raw_ots: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "on-time-shipment.metrics-table-columns"))).text.split("\n")
                ots: str = raw_ots[1]
                ots_units: str = raw_ots[2].split(" ")[0]

                ots_row = _metrics_rows[f"sfp_{size_key}_ots"]
                sh_metrics.range(f"{col}{ots_row}").value = ots
                sh_metrics.range(f"{col}{ots_row + 1}").value = ots_units

                time.sleep(5)

        driver.quit()

        curr_month: int = datetime.now().month
        curr_date: str = datetime.now().strftime("%d/%Y")
        date_str: str = f"{curr_month}/{curr_date}"

        log.info("Organizing charts, saving and closing workbook.")
        _run_checked_macro(excel, ah_wb, "modUtilities.resizeCharts")
        time.sleep(3)
        sh_metrics.range("B1").value = date_str
        sh_dash.range("B1").value = date_str
        ah_wb.save()
        ah_wb.close()
        _end_excel(excel)
        excel = None

        log.info("Sending email.")
        time.sleep(30)
        outlook.send_email(
            account=sender_email,
            subject=f"Account Health Metrics - {date_str}",
            body=_email_body(),
            to=to_email,
            cc=cc_email,
            attachments=[ah_wb_path],
            show=True,
            send=True
        )
        log.info("Email sent.")

    except (KeyboardInterrupt, SystemExit):
        log.warning("Script interrupted by user.")
        raise SystemExit(0)

    except Exception:
        alert_utils.handle_crash(driver, traceback.format_exc(), "Account Health")
        raise SystemExit(1)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if excel is not None:
            _end_excel(excel)
        if excel_pin is not None:
            excel_pin.Close()


if ask_user("Run now?", "Amazon Account Health"):
    main()
run_on_schedule(main, hour=11, minute=0, day_of_week="mon-fri")
