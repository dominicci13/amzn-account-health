import os
import io
import json
import time
import ctypes
import traceback
import pandas as pd
import xlwings as xw
from PIL import Image
import win32clipboard
from rich import print
from dotenv import load_dotenv
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, custom_functions, accounts, outlook
from selenium.common.exceptions import SessionNotCreatedException, TimeoutException

###############################################################################################################################################
#Get the user and working directory
directory: str = os.getcwd()
win_user: str = os.getlogin()

#Get Seller Central credentials from environment
load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")
sender_email: str = os.getenv("SENDER_EMAIL")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]

#Load the Excel cell address config
with open(f"{directory}/config/cells.json") as f:
    cells = json.load(f)

#Set Chrome User Data Directory
user_data_dir: str = f"C:/ChromeAutomationProfile"

#Get Amazon accounts links
Accounts = accounts.Amazon()

##################################################################################################################################################
# Create the body of the email
body = """
Good morning,<br><br>
Please find attached Account Health Metrics file updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""

##################################################################################################################################################
def seconds_until_target(TargetTime: str):
    #Calculate the number of seconds until the target time
    now: datetime = datetime.now()
    TargetTime = datetime.strptime(TargetTime, "%H:%M:%S").replace(year=now.year, month=now.month, day=now.day)

    if TargetTime < now:
        TargetTime += timedelta(days=1)

    return (TargetTime - now).total_seconds()

##################################################################################################################################################
def ShouldRun() -> bool:
    #Check if today is not Saturday or Sunday
    today: str = datetime.now().strftime("%A")

    return today not in ["Saturday", "Sunday"]

##################################################################################################################################################
#Ask the user if they want to start the process now
BtnPressed = ctypes.windll.user32.MessageBoxW(
    0,
    "Do you want to start the script now?",
    "Account Health Metrics",
    4 | 0x20
)

while True:
    #Time to start
    StartTime = "11:00:00"
    StartHour = int(StartTime.split(":")[0])
    StartMin = StartTime.split(":")[1]

    if ShouldRun():
        SleepTime = seconds_until_target(StartTime)
        nowHour = int(datetime.now().strftime("%H"))
        today: str = datetime.now().strftime("%A")
        tomorrow: str = custom_functions.tomorrow()

        #If the user pressed "Yes", then start the process
        if BtnPressed == 7:
            if tomorrow in ["Saturday", "Sunday"]:
                if nowHour >= StartHour:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated on Monday at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated on Monday at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} AM.")

            else:
                if nowHour >= StartHour:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated tomorrow {tomorrow} at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated tomorrow {tomorrow} at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"[cyan][INFO][/cyan] Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} AM.")

            #Sleep until just before the Start time
            time.sleep(max(SleepTime - 1, 0))

            #Loop to ensure that we catch the exact time
            while datetime.now().strftime("%H:%M:%S") != StartTime:
                time.sleep(0.5)

            if tomorrow in ["Saturday", "Sunday"] and nowHour >= StartHour:
                continue

        #Reset the value of the button
        BtnPressed = 7

        ###############################################################################################################################################
        #Set workbook properties
        AHwb: str = f"{directory}/Amazon/Reports/AH-Metrics.xlsm"

        print("[cyan][INFO][/cyan] Opening workbook and removing charts.")
        AHMetrics = xw.Book(AHwb)
        shMetrics = AHMetrics.sheets(1)
        shDash = AHMetrics.sheets(2)
        delChar = AHMetrics.macro("Module1.DeleteChar")
        resizeChar = AHMetrics.macro("Module1.ResizeChar")

        delChar() #Run Macro

        ###############################################################################################################################################
        #Initialize Chrome
        opening_browser = True
        while opening_browser:
            try:
                driver: object = chrome.start_browser(
                    user_data_dir,
                    "Default",
                    headless=True
                )
                opening_browser = False
            except (SessionNotCreatedException, RuntimeError):
                print("[bold red][ERROR][/bold red] Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
                custom_functions.kill_app("chrome")
                time.sleep(5)

        #Navigate through each account
        for account, url in Accounts.items():
            match account:
                case "FocusCam":
                    root = "SellerOrg Corp"
                case "LifeS":
                    root = "Lifestyle By Focus"
                case "XtraB":
                    root = "XtraBargains"
                case "KnoxGear":
                    root = "Knox Gear"
                case "Apple":
                    root = "Apple Renewed Focus"
                case "FocusHome":
                    root = "Focus Home"

            col = cells["metrics_columns"][account]
            dash = cells["dashboard"][account]

            print(f"[cyan][INFO][/cyan] Navigating to [cyan]{root}[/cyan] account.")
            driver.get(url)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.Amazon_login(driver, username, password)

                    if not code:
                        print("[bold red][ERROR][/bold red] Failed to log in to Amazon. Trying again.")
                        driver.get(url)
                        driver.switch_to_window(0)

            except TimeoutException:
                pass

            ###############################################################################################################################################
            trying = True
            while trying:
                try:
                    #Navigate to Account Health Dashboard and get Shipping Performance numbers
                    driver.get("https://sellercentral.amazon.com/performance/dashboard")
                    driver.switch_to_window(0)

                    print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] Account Health Statistics.")
                    ShipRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-late-shipment-rate-row"))).text.split("\n")
                    PreCancelRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-cancellation-rate-row"))).text.split("\n")
                    ValidTrackRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-view-tracking-rate-row"))).text.split("\n")
                    trying = False

                except TimeoutException:
                    print("[bold red][ERROR][/bold red] Failed to get [cyan]Account Health[/cyan] Statistics. Retrying.")
                    time.sleep(5)

            try:
                LSR: str = ShipRate[2]
                LSROrders: str = ShipRate[3].replace("orders", "orders (30 days)")
            except IndexError:
                LSR = "N/A"
                LSROrders = "N/A"

            try:
                PCR: str = PreCancelRate[2]
                PCROrders: str = PreCancelRate[3].replace("orders", "orders (7 days)")
            except IndexError:
                PCR = "N/A"
                PCROrders = "N/A"

            try:
                VTR: str = ValidTrackRate[2]
                VTROrders: str = ValidTrackRate[3].replace("orders", "orders (30 days)")
            except IndexError:
                VTR = "N/A"
                VTROrders = "N/A"

            values: list[str] = [LSR, LSROrders, PCR, PCROrders, VTR, VTROrders]
            num_columns = 1
            rows: list[list[str]] = [values[i:i+num_columns] for i in range(0, len(values), num_columns)]
            df = pd.DataFrame(rows[0:])

            shMetrics.range(f"{col}{cells['metrics_rows']['account_health']}").value = df.values

            ###############################################################################################################################################
            #Navigate to Prime Performance and get the data
            driver.get("https://sellercentral.amazon.com/performance/eligibilities?ref=sp-st-dash-mons-elgibl")
            driver.switch_to_window(0)

            print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] Prime Performance Statistics.")
            PrimePerformance: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "guaranteed-delivery"))).text.split("\n")

            AccStatus: str = PrimePerformance[1]

            if AccStatus.startswith("✓"):
                AccStatus = "Eligible"
            else:
                AccStatus = "Not Eligible"

            OTDR: str = PrimePerformance[3]
            OTDROrders: str = PrimePerformance[5]
            PFCR: str = PrimePerformance[7]
            PFCROrders: str = PrimePerformance[9]
            VTUR: str = PrimePerformance[11]
            VTUROrders: str = PrimePerformance[13]

            values = [AccStatus, OTDR, OTDROrders, PFCR, PFCROrders, VTUR, VTUROrders]
            num_columns = 1
            rows = [values[i:i+num_columns] for i in range(0, len(values), num_columns)]
            df = pd.DataFrame(rows[0:])

            shMetrics.range(f"{col}{cells['metrics_rows']['prime_performance']}").value = df.values

            ###############################################################################################################################################
            #Create a dictionary of the Seller Fulfilled Prime performance URLs
            sfp_performance_urls = {
                "Standard-size": "https://sellercentral.amazon.com/seller-fulfilled-prime/seller-performance?tierName=STD",
                "Oversize": "https://sellercentral.amazon.com/seller-fulfilled-prime/seller-performance?tierName=OS"
            }

            #Navigate to Seller Fulfilled Prime performance
            for size, sfp_url in sfp_performance_urls.items():
                driver.get(sfp_url)
                driver.switch_to_window(0)
                time.sleep(3)

                size_key = "standard" if size == "Standard-size" else "oversize"

                print(f"[cyan][INFO][/cyan] Checking [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Program current status.")
                ProgramStatus: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "a-alert-content"))).text
                time.sleep(2)

                if ProgramStatus == "In Trial":
                    TrialEndDate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sfp-trial-performance-end-date"]'))).text.split(", ")
                    ProgramStatus = f"{ProgramStatus} (until {TrialEndDate[0]})"
                elif ProgramStatus.startswith("Your Seller Fulfilled Prime performance has failed "):
                    ProgramStatus = "Revoked"

                shMetrics.range(f"{col}{cells['metrics_rows'][f'sfp_{size_key}_status']}").value = ProgramStatus

                ###############################################################################################################################################
                #Show past seven days Speed metrics
                print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Speed metric charts.")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "show-gvs-metrics-by-date-filter-past-seven-days"))).click()
                element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "a-list-item")))
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(3)

                ###############################################################################################################################################
                #Build the CSS Selector for the metrics chart depending on the size
                if size == "Standard-size":
                    chart_css = "#std-domestic-delivery-speed-graph > div > div:nth-child(1)"
                elif size == "Oversize":
                    chart_css = "#os-domestic-delivery-speed-graph > div > div:nth-child(1)"

                #Copy the chart element to clipboard, and paste it into Excel
                Graph = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    chart_css
                )))

                Location = Graph.location
                Size = Graph.size

                png = driver.get_screenshot_as_png()
                Char = Image.open(io.BytesIO(png))

                #Change chart "y" position depending if banner "In Trial" is there
                left = Location['x']
                top = Location['y'] - 145
                right = Location['x'] + Size['width']
                bottom = Location['y'] + Size['height'] - 130

                Char = Char.crop((left, top, right, bottom))

                #Save chart screenshot to clipboard
                Output = io.BytesIO()
                Char.convert("RGB").save(Output, "BMP")
                Data = Output.getvalue()[14:]
                Output.close()
                custom_functions.send_to_clipboard(win32clipboard.CF_DIB, Data)

                try:
                    custom_functions.paste_image_from_clipboard(shDash, dash[f"{size_key}_chart_cell"])
                except Exception:
                    print("[bold red][ERROR][/bold red] Failed to paste chart into Excel")

                ###############################################################################################################################################
                #Build the ID for the metrics table depending on the size
                if size == "Standard-size":
                    table_id = "std-domestic-delivery-speed-details-table"
                elif size == "Oversize":
                    table_id = "os-domestic-delivery-speed-details-table"

                #Current size metrics
                print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Speed metrics Statistics.")
                RawMetrics: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.ID,
                    table_id
                ))).text.split("\n")

                OneDay: list[str] = RawMetrics[-3].split(" ")
                PercOneDay: str = OneDay[3]
                ViewsOneDay: str = OneDay[-1]

                TwoDays: list[str] = RawMetrics[-2].split(" ")
                PercTwoDays: str = TwoDays[3]
                ViewsTwoDays: str = TwoDays[-1]

                TwoDaysPlus: list[str] = RawMetrics[-1].split(" ")
                PercTwoDaysPlus: str = TwoDaysPlus[3]
                ViewsTwoDaysPlus: str = TwoDaysPlus[-1]

                speed_col = dash[f"{size_key}_speed_col"]
                speed_row = dash[f"{size_key}_speed_row"]
                shDash.range(f"{speed_col}{speed_row}").value = [PercOneDay, ViewsOneDay]
                shDash.range(f"{speed_col}{speed_row + 2}").value = [PercTwoDays, ViewsTwoDays]
                shDash.range(f"{speed_col}{speed_row + 4}").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                ###############################################################################################################################################
                #Show past seven days Fulfillment metrics
                print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Fulfillment Statistics.")
                FulfillmentButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.ID,
                    "show-fulfillment-metrics-by-date-filter-past-seven-days"
                )))
                driver.execute_script("arguments[0].scrollIntoView(true);", FulfillmentButton)
                FulfillmentButton.click()
                time.sleep(3)

                #On-time delivery metrics
                RawOTD: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "on-time-delivery.metrics-table-columns"
                ))).text.split("\n")[1].split(" ")
                OTD: str = RawOTD[0]
                OTDUnits: str = RawOTD[4]

                otd_row = cells["metrics_rows"][f"sfp_{size_key}_otd"]
                shMetrics.range(f"{col}{otd_row}").value = OTD
                shMetrics.range(f"{col}{otd_row + 1}").value = OTDUnits

                #Valid tracking rate metrics
                RawVTR: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "valid-tracking.metrics-table-columns"
                ))).text.split("\n")
                VTR = RawVTR[1]
                VTRPackages: str = RawVTR[3].split(" ")[0]

                vtr_row = cells["metrics_rows"][f"sfp_{size_key}_vtr"]
                shMetrics.range(f"{col}{vtr_row}").value = VTR
                shMetrics.range(f"{col}{vtr_row + 1}").value = VTRPackages

                #Cancellation rate metrics
                RawCR: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "cancellation.metrics-table-columns"
                ))).text.split("\n")
                CR: str = RawCR[1]
                CRUnits: str = RawCR[3].split(" ")[0]

                cr_row = cells["metrics_rows"][f"sfp_{size_key}_cr"]
                shMetrics.range(f"{col}{cr_row}").value = CR
                shMetrics.range(f"{col}{cr_row + 1}").value = CRUnits

                ###############################################################################################################################################
                #Show past seven days Supporting metrics
                print(f"[cyan][INFO][/cyan] Getting [cyan]{root}[/cyan] - [cyan]{size}[/cyan] Supporting Statistics.")
                SupportButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.ID,
                    "show-supporting-metrics-by-date-filter-past-seven-days"
                )))
                driver.execute_script("arguments[0].scrollIntoView(true);", SupportButton)
                SupportButton.click()
                time.sleep(3)

                #On-time shipment metrics
                RawOTS: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "on-time-shipment.metrics-table-columns"
                ))).text.split("\n")
                OTS: str = RawOTS[1]
                OTSUnits: str = RawOTS[2].split(" ")[0]

                ots_row = cells["metrics_rows"][f"sfp_{size_key}_ots"]
                shMetrics.range(f"{col}{ots_row}").value = OTS
                shMetrics.range(f"{col}{ots_row + 1}").value = OTSUnits

                time.sleep(5)

        #Close browser and save workbook
        driver.quit()

        #Set all necessary date variables
        currDate: str = datetime.now().strftime("%d/%Y")
        currMonth: int = datetime.now().month
        Date: str = f"{currMonth}/{currDate}"

        #Call macros
        print("[cyan][INFO][/cyan] Organizing charts, saving and closing workbook.")
        resizeChar()
        time.sleep(3)
        shMetrics.range("B1").value = Date
        shDash.range("B1").value = Date
        AHMetrics.save()
        AHMetrics.close()

        print("[cyan][INFO][/cyan] Loading workbook and sending email.")
        time.sleep(30)
        outlook.send_email(
            account=sender_email,
            subject=f"Account Health Metrics - {Date}",
            body=body,
            to=to_email,
            cc=cc_email,
            attachments=[AHwb],
            show=True,
            send=True
        )

        print("[cyan][INFO][/cyan] Email has been sent.")

    #Sleep 60 seconds before starting over
    time.sleep(60)
