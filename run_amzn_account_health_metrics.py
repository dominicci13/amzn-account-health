import os
import io
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

#Set Chrome User Data Directory
#user_data_dir: str = f"C:/Users/{win_user}/AppData/Local/Google/Chrome/User Data"
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
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated on Monday at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated on Monday at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} AM.")

            else:
                if nowHour >= StartHour:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated tomorrow {tomorrow} at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated tomorrow {tomorrow} at {StartHour}:{StartMin} AM.")
                else:
                    if StartHour > 12:
                        StartHour -= 12
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} PM.")
                    else:
                        print(f"'[INFO]' Account Health Metrics dashboard will be updated today at {StartHour}:{StartMin} AM.")

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

        print("'[INFO]' Opening workbook and removing charts.")
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
                print("'[ERROR]' Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
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

            print(f"'[INFO]' Navigating to '{root}' account.")
            driver.get(url)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.Amazon_login(driver, username, password)

                    if not code:
                        print("'[ERROR]' Failed to log in to Amazon. Trying again.")
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

                    print(f"'[INFO]' Getting '{root}' Account Health Statistics.")
                    ShipRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-late-shipment-rate-row"))).text.split("\n")
                    PreCancelRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-cancellation-rate-row"))).text.split("\n")
                    ValidTrackRate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "shipping-view-tracking-rate-row"))).text.split("\n")
                    trying = False

                except TimeoutException:
                    print("'[ERROR]' Failed to get 'Account Health' Statistics. Retrying.")
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
            rows: list[list[str]] = [values[i:i+num_columns] for i in range(0, len(values), num_columns)] #Split raw data into rows
            df = pd.DataFrame(rows[0:])

            #Write retrieved data into spreadsheet cells accordingly
            if account == "LifeS":
                shMetrics.range("C7").value = df.values
            elif account == "FocusCam":
                shMetrics.range("D7").value = df.values
            elif account == "XtraB":
                shMetrics.range("E7").value = df.values
            elif account == "KnoxGear":
                shMetrics.range("F7").value = df.values
            elif account == "Apple":
                shMetrics.range("G7").value = df.values
            elif account == "FocusHome":
                shMetrics.range("H7").value = df.values

            ###############################################################################################################################################
            #Navigate to Prime Performance and get the data
            driver.get("https://sellercentral.amazon.com/performance/eligibilities?ref=sp-st-dash-mons-elgibl")
            driver.switch_to_window(0)

            print(f"'[INFO]' Getting '{root}' Prime Performance Statistics.")
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
            rows = [values[i:i+num_columns] for i in range(0, len(values), num_columns)] #Split raw data into rows
            df = pd.DataFrame(rows[0:])

            #Write retrieved data into spreadsheet cells accordingly
            if account == "LifeS":
                shMetrics.range("C17").value = df.values
            elif account == "FocusCam":
                shMetrics.range("D17").value = df.values
            elif account == "XtraB":
                shMetrics.range("E17").value = df.values
            elif account == "KnoxGear":
                shMetrics.range("F17").value = df.values
            elif account == "Apple":
                shMetrics.range("G17").value = df.values
            elif account == "FocusHome":
                shMetrics.range("H17").value = df.values

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

                print(f"'[INFO]' Checking '{root}' - '{size}' Program current status.")
                ProgramStatus: str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "a-alert-content"))).text
                time.sleep(2)

                if ProgramStatus == "In Trial":
                    TrialEndDate: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sfp-trial-performance-end-date"]'))).text.split(", ")
                    ProgramStatus = f"{ProgramStatus} (until {TrialEndDate[0]})"
                elif ProgramStatus.startswith("Your Seller Fulfilled Prime performance has failed "):
                    ProgramStatus = "Revoked"

                #Write retrieved data into spreadsheet cells accordingly
                if size == "Standard-size":
                    if account == "LifeS":
                        shMetrics.range("C28").value = ProgramStatus
                    elif account == "FocusCam":
                        shMetrics.range("D28").value = ProgramStatus
                    elif account == "XtraB":
                        shMetrics.range("E28").value = ProgramStatus
                    elif account == "KnoxGear":
                        shMetrics.range("F28").value = ProgramStatus
                    elif account == "Apple":
                        shMetrics.range("G28").value = ProgramStatus
                    elif account == "FocusHome":
                        shMetrics.range("H28").value = ProgramStatus

                elif size == "Oversize":
                    if account == "LifeS":
                        shMetrics.range("C38").value = ProgramStatus
                    elif account == "FocusCam":
                        shMetrics.range("D38").value = ProgramStatus
                    elif account == "XtraB":
                        shMetrics.range("E38").value = ProgramStatus
                    elif account == "KnoxGear":
                        shMetrics.range("F38").value = ProgramStatus
                    elif account == "Apple":
                        shMetrics.range("G38").value = ProgramStatus
                    elif account == "FocusHome":
                        shMetrics.range("H38").value = ProgramStatus

                ###############################################################################################################################################
                #Show past seven days Speed metrics
                print(f"'[INFO]' Getting '{root}' - '{size}' Speed metric charts.")
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
                if ProgramStatus.startswith("In Trial"):
                    left = Location['x']
                    top = Location['y'] - 145
                    right = Location['x'] + Size['width']
                    bottom = Location['y'] + Size['height'] - 130
                else:
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
                    #Write the chart to Excel according to the account
                    if size == "Standard-size":
                        if account == "LifeS":
                            custom_functions.paste_image_from_clipboard(shDash, "F10")
                        elif account == "FocusCam":
                            custom_functions.paste_image_from_clipboard(shDash, "M10")
                        elif account == "XtraB":
                            custom_functions.paste_image_from_clipboard(shDash, "F29")
                        elif account == "KnoxGear":
                            custom_functions.paste_image_from_clipboard(shDash, "M29")
                        elif account == "Apple":
                            custom_functions.paste_image_from_clipboard(shDash, "F48")
                        elif account == "FocusHome":
                            custom_functions.paste_image_from_clipboard(shDash, "M48")

                    elif size == "Oversize":
                        if account == "LifeS":
                            custom_functions.paste_image_from_clipboard(shDash, "F18")
                        elif account == "FocusCam":
                            custom_functions.paste_image_from_clipboard(shDash, "M18")
                        elif account == "XtraB":
                            custom_functions.paste_image_from_clipboard(shDash, "F37")
                        elif account == "KnoxGear":
                            custom_functions.paste_image_from_clipboard(shDash, "M37")
                        elif account == "Apple":
                            custom_functions.paste_image_from_clipboard(shDash, "F56")
                        elif account == "FocusHome":
                            custom_functions.paste_image_from_clipboard(shDash, "M56")
                except:
                    print("'[ERROR]' Failed to paste chart into Excel")

                ###############################################################################################################################################
                #Build the ID for the metrics table depending on the size
                if size == "Standard-size":
                    table_id = "std-domestic-delivery-speed-details-table"
                elif size == "Oversize":
                    table_id = "os-domestic-delivery-speed-details-table"

                #Current size metrics
                print(f"'[INFO]' Getting '{root}' - '{size}' Speed metrics Statistics.")
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

                #Write the chart to Excel according to the account
                if size == "Standard-size":
                    if account == "LifeS":
                        shDash.range("C10").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C12").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C14").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "FocusCam":
                        shDash.range("J10").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J12").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J14").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "XtraB":
                        shDash.range("C29").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C31").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C33").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "KnoxGear":
                        shDash.range("J29").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J31").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J33").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "Apple":
                        shDash.range("C48").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C50").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C52").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "FocusHome":
                        shDash.range("J48").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J50").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J52").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                elif size == "Oversize":
                    if account == "LifeS":
                        shDash.range("C18").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C20").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C22").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "FocusCam":
                        shDash.range("J18").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J20").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J22").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "XtraB":
                        shDash.range("C37").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C39").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C41").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "KnoxGear":
                        shDash.range("J37").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J39").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J41").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "Apple":
                        shDash.range("C56").value = [PercOneDay, ViewsOneDay]
                        shDash.range("C58").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("C60").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                    elif account == "FocusHome":
                        shDash.range("J56").value = [PercOneDay, ViewsOneDay]
                        shDash.range("J58").value = [PercTwoDays, ViewsTwoDays]
                        shDash.range("J60").value = [PercTwoDaysPlus, ViewsTwoDaysPlus]

                ###############################################################################################################################################
                #Show past seven days Fulfillment metrics
                print(f"'[INFO]' Getting '{root}' - '{size}' Fulfillment Statistics.")
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

                #Write retrieved data into spreadsheet cells accordingly
                if size == "Standard-size":
                    if account == "LifeS":
                        shMetrics.range("C31").value = OTD
                        shMetrics.range("C32").value = OTDUnits

                    elif account == "FocusCam":
                        shMetrics.range("D31").value = OTD
                        shMetrics.range("D32").value = OTDUnits

                    elif account == "XtraB":
                        shMetrics.range("E31").value = OTD
                        shMetrics.range("E32").value = OTDUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F31").value = OTD
                        shMetrics.range("F32").value = OTDUnits

                    elif account == "Apple":
                        shMetrics.range("G31").value = OTD
                        shMetrics.range("G32").value = OTDUnits

                    elif account == "FocusHome":
                        shMetrics.range("H31").value = OTD
                        shMetrics.range("H32").value = OTDUnits

                elif size == "Oversize":
                    if account == "LifeS":
                        shMetrics.range("C41").value = OTD
                        shMetrics.range("C42").value = OTDUnits

                    elif account == "FocusCam":
                        shMetrics.range("D41").value = OTD
                        shMetrics.range("D42").value = OTDUnits

                    elif account == "XtraB":
                        shMetrics.range("E41").value = OTD
                        shMetrics.range("E42").value = OTDUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F41").value = OTD
                        shMetrics.range("F42").value = OTDUnits

                    elif account == "Apple":
                        shMetrics.range("G41").value = OTD
                        shMetrics.range("G42").value = OTDUnits

                    elif account == "FocusHome":
                        shMetrics.range("H41").value = OTD
                        shMetrics.range("H42").value = OTDUnits

                #Valid tracking rate metrics
                RawVTR: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "valid-tracking.metrics-table-columns"
                ))).text.split("\n")
                VTR = RawVTR[1]
                VTRPackages: str = RawVTR[3].split(" ")[0]

                #Write retrieved data into spreadsheet cells accordingly
                if size == "Standard-size":
                    if account == "LifeS":
                        shMetrics.range("C33").value = VTR
                        shMetrics.range("C34").value = VTRPackages

                    elif account == "FocusCam":
                        shMetrics.range("D33").value = VTR
                        shMetrics.range("D34").value = VTRPackages

                    elif account == "XtraB":
                        shMetrics.range("E33").value = VTR
                        shMetrics.range("E34").value = VTRPackages

                    elif account == "KnoxGear":
                        shMetrics.range("F33").value = VTR
                        shMetrics.range("F34").value = VTRPackages

                    elif account == "Apple":
                        shMetrics.range("G33").value = VTR
                        shMetrics.range("G34").value = VTRPackages

                    elif account == "FocusHome":
                        shMetrics.range("H33").value = VTR
                        shMetrics.range("H34").value = VTRPackages

                elif size == "Oversize":
                    if account == "LifeS":
                        shMetrics.range("C43").value = VTR
                        shMetrics.range("C44").value = VTRPackages

                    elif account == "FocusCam":
                        shMetrics.range("D43").value = VTR
                        shMetrics.range("D44").value = VTRPackages

                    elif account == "XtraB":
                        shMetrics.range("E43").value = VTR
                        shMetrics.range("E44").value = VTRPackages

                    elif account == "KnoxGear":
                        shMetrics.range("F43").value = VTR
                        shMetrics.range("F44").value = VTRPackages

                    elif account == "Apple":
                        shMetrics.range("G43").value = VTR
                        shMetrics.range("G44").value = VTRPackages

                    elif account == "FocusHome":
                        shMetrics.range("H43").value = VTR
                        shMetrics.range("H44").value = VTRPackages

                #Cancellation rate metrics
                RawCR: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((
                    By.CLASS_NAME,
                    "cancellation.metrics-table-columns"
                ))).text.split("\n")
                CR: str = RawCR[1]
                CRUnits: str = RawCR[3].split(" ")[0]

                #Write retrieved data into spreadsheet cells accordingly
                if size == "Standard-size":
                    if account == "LifeS":
                        shMetrics.range("C35").value = CR
                        shMetrics.range("C36").value = CRUnits

                    elif account == "FocusCam":
                        shMetrics.range("D35").value = CR
                        shMetrics.range("D36").value = CRUnits

                    elif account == "XtraB":
                        shMetrics.range("E35").value = CR
                        shMetrics.range("E36").value = CRUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F35").value = CR
                        shMetrics.range("F36").value = CRUnits

                    elif account == "Apple":
                        shMetrics.range("G35").value = CR
                        shMetrics.range("G36").value = CRUnits

                    elif account == "FocusHome":
                        shMetrics.range("H35").value = CR
                        shMetrics.range("H36").value = CRUnits

                elif size == "Oversize":
                    if account == "LifeS":
                        shMetrics.range("C45").value = CR
                        shMetrics.range("C46").value = CRUnits

                    elif account == "FocusCam":
                        shMetrics.range("D45").value = CR
                        shMetrics.range("D46").value = CRUnits

                    elif account == "XtraB":
                        shMetrics.range("E45").value = CR
                        shMetrics.range("E46").value = CRUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F45").value = CR
                        shMetrics.range("F46").value = CRUnits

                    elif account == "Apple":
                        shMetrics.range("G45").value = CR
                        shMetrics.range("G46").value = CRUnits

                    elif account == "FocusHome":
                        shMetrics.range("H45").value = CR
                        shMetrics.range("H46").value = CRUnits

                ###############################################################################################################################################
                #Show past seven days Supporting metrics
                print(f"'[INFO]' Getting '{root}' - '{size}' Supporting Statistics.")
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

                #Write retrieved data into spreadsheet cells accordingly
                if size == "Standard-size":
                    if account == "LifeS":
                        shMetrics.range("C29").value = OTS
                        shMetrics.range("C30").value = OTSUnits

                    elif account == "FocusCam":
                        shMetrics.range("D29").value = OTS
                        shMetrics.range("D30").value = OTSUnits

                    elif account == "XtraB":
                        shMetrics.range("E29").value = OTS
                        shMetrics.range("E30").value = OTSUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F29").value = OTS
                        shMetrics.range("F30").value = OTSUnits

                    elif account == "Apple":
                        shMetrics.range("G29").value = OTS
                        shMetrics.range("G30").value = OTSUnits
                        
                    elif account == "FocusHome":
                        shMetrics.range("H29").value = OTS
                        shMetrics.range("H30").value = OTSUnits

                elif size == "Oversize":
                    if account == "LifeS":
                        shMetrics.range("C39").value = OTS
                        shMetrics.range("C40").value = OTSUnits

                    elif account == "FocusCam":
                        shMetrics.range("D39").value = OTS
                        shMetrics.range("D40").value = OTSUnits

                    elif account == "XtraB":
                        shMetrics.range("E39").value = OTS
                        shMetrics.range("E40").value = OTSUnits

                    elif account == "KnoxGear":
                        shMetrics.range("F39").value = OTS
                        shMetrics.range("F40").value = OTSUnits

                    elif account == "Apple":
                        shMetrics.range("G39").value = OTS
                        shMetrics.range("G40").value = OTSUnits

                    elif account == "FocusHome":
                        shMetrics.range("H39").value = OTS
                        shMetrics.range("H40").value = OTSUnits

                time.sleep(5)

        #Close Firefox and save workbook
        driver.quit()

        #Set all necessary date variables
        currDate: str = datetime.now().strftime("%d/%Y")
        currMonth: int = datetime.now().month
        Date: str = f"{currMonth}/{currDate}"

        #Call macros
        print("'[INFO]' Organizing charts, saving and closing workbook.")
        resizeChar()
        time.sleep(3)
        shMetrics.range("B1").value = Date
        shDash.range("B1").value = Date
        AHMetrics.save()
        AHMetrics.close()

        print("'[INFO]' Loading workbook and sending email.")
        time.sleep(30)
        outlook.send_email(
            account="user@example.com",
            subject=f"Account Health Metrics - {Date}",
            body=body,
            to=["user@example.com"],
            cc=["user@example.com", "user@example.com"],
            attachments=[AHwb],
            show=True,
            send=True
        )

        #Save and close workbook
        print("'[INFO]' Email has been sent.")

    #Sleep 60 seconds before starting over
    time.sleep(60)