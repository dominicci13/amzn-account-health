import io
import os
import time
import traceback
from pathlib import Path
import pandas as pd
import xlwings as xw
from PIL import Image
import win32clipboard
from dotenv import load_dotenv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, custom_functions, accounts, outlook, alert_utils
from fc_utils.config_utils import get_env, load_config_safe
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.ui_utils import ask_user
from fc_utils.logging_utils import setup_logging
from selenium.common.exceptions import TimeoutException


log = setup_logging("amzn_account_health_metrics")
load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")
sender_email: str = os.getenv("SENDER_EMAIL", "")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)

_accounts_cfg = load_config_safe(Path.cwd() / "config" / "accounts.json")
_metrics_columns: dict[str, str] = _accounts_cfg.get("account_health_metrics_columns", {})
_dashboard: dict[str, dict] = _accounts_cfg.get("account_health_dashboard", {})

_layout = load_config_safe(Path.cwd() / "config" / "metrics_layout.json")
_metrics_rows: dict[str, int] = _layout.get("metrics_rows", {})

_paths = load_config_safe(Path.cwd() / "config" / "paths.json")
ah_wb_path: str = _paths["ah_wb_path"]

body = """
Good morning,<br><br>
Please find attached Account Health Metrics file updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""


def main() -> None:
    """Scrape account health metrics from each Amazon account and email the workbook.

    Opens the AH-Metrics workbook, loops through each Amazon account to scrape
    shipping performance, Prime eligibility, and Seller Fulfilled Prime stats,
    pastes charts into the dashboard sheet, saves the workbook, and emails it.
    """
    driver = None
    excel = None
    try:
        log.info("Opening workbook and removing charts.")
        excel = xw.App(visible=False)
        excel.display_alerts = False
        excel.screen_updating = False
        ah_wb = excel.books.open(ah_wb_path)
        sh_metrics = ah_wb.sheets(1)
        sh_dash = ah_wb.sheets(2)
        del_char = ah_wb.macro("Module1.DeleteChar")
        resize_char = ah_wb.macro("Module1.ResizeChar")

        del_char()

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

            # Prime performance stats
            driver.get("https://sellercentral.amazon.com/performance/eligibilities?ref=sp-st-dash-mons-elgibl")
            driver.switch_to_window(0)

            log.info(f"Getting [cyan]{root}[/cyan] Prime Performance Statistics.")
            prime_perf: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "guaranteed-delivery"))).text.split("\n")

            acc_status: str = "Eligible" if prime_perf[1].startswith("✓") else "Not Eligible"
            otdr: str = prime_perf[3]
            otdr_orders: str = prime_perf[5]
            pfcr: str = prime_perf[7]
            pfcr_orders: str = prime_perf[9]
            vtur: str = prime_perf[11]
            vtur_orders: str = prime_perf[13]

            values = [acc_status, otdr, otdr_orders, pfcr, pfcr_orders, vtur, vtur_orders]
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

                location = graph.location
                element_size = graph.size

                png = driver.get_screenshot_as_png()
                chart_img = Image.open(io.BytesIO(png))

                left = location['x']
                top = location['y'] - 145
                right = location['x'] + element_size['width']
                bottom = location['y'] + element_size['height'] - 130

                chart_img = chart_img.crop((left, top, right, bottom))

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
        resize_char()
        time.sleep(3)
        sh_metrics.range("B1").value = date_str
        sh_dash.range("B1").value = date_str
        ah_wb.save()
        ah_wb.close()
        excel.quit()
        excel = None

        log.info("Sending email.")
        time.sleep(30)
        outlook.send_email(
            account=sender_email,
            subject=f"Account Health Metrics - {date_str}",
            body=body,
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
        alert_utils.handle_crash(driver, traceback.format_exc(), "Account Health Metrics")
        raise SystemExit(1)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if excel is not None:
            try:
                excel.quit()
            except Exception:
                pass


if ask_user("Run now?", "Amazon Account Health Metrics"):
    main()
run_on_schedule(main, hour=11, minute=0, day_of_week="mon-fri")
