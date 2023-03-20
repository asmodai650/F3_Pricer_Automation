import argparse
import datetime
import time
import traceback
import pymssql
import pyautogui as pa
import uuid
import getpass
import os
import __main__

from datetime import datetime
from sikuliWrapper import *
from ScrapeQManager import scrapeQ


#region ARGUMENTS

parser = argparse.ArgumentParser()
parser.add_argument('--sqlu', dest='SQLUSER')
parser.add_argument('--sqlp', dest='SQLPASS')
parser.add_argument('--u0', dest='CitrixUser')
parser.add_argument('--p0', dest='CitrixPass')
parser.add_argument('--q', dest='scrapeq')
parser.add_argument('--msid', dest='userMS')
parser.add_argument('--uuid', dest='uuid')
args, unknown = parser.parse_known_args()

if args.SQLUSER is None:
    scriptuuid = str(uuid.uuid4()) #just a random unique number to identify which scrape was running
    msid = input('Enter MS ID: ')
    UserMSID = msid 
    password = getpass.getpass("Enter MS password (cursor won't move while typing): ")
    msid = f'MS\\{msid}'
    username = msid
    scrapeQName = 'OptumCare Facets F3 Pricer'
else:
    username = args.SQLUSER
    password = args.SQLPASS
    scriptuuid = args.uuid
    scrapeQName = args.scrapeq
    userMS = args.userMS
    UserMSID = args.userMS
    if 'ms\\' not in username.lower():
        msid = f'MS\\{username}'
    else:
        msid = username

#endregion ARGUMENTS

#region VARIABLES

mainfolder = os.path.dirname(os.path.abspath(os.path.join(os.getcwd(), __main__.__file__)))
imgLocation = os.path.join(mainfolder,'scrapeimages')

projectid = '1394'

#endregion VARIABLES

#region FUNCTIONS

def setupFacets():
    print("Trying to maximize screen")
    time.sleep(3)
    
    # Check to see if Facets is ready to go
    if pa.locateOnScreen(os.path.join(imgLocation, 'alreadyInFacets.png'), confidence=.7) is not None:
        print(f"Facets is already at the Hospital Claims tab .")
        pa.moveTo(x = 430, y = 117)
        time.sleep(.5)

        alwaysVisible()

        return

    # Maximize Facets Screen
    if pa.locateOnScreen(os.path.join(imgLocation, 'FacetsTopBar.png'), confidence=.7) is not None: 
        print("Found FacetsTopBar .")
        rightClick('FacetsTopBar.png')
        time.sleep(0.5)
        pa.press('up')
        time.sleep(0.5)
        pa.press('up')
        time.sleep(0.5)
        pa.press('enter')
        time.sleep(1)

        alwaysVisible()
    else:
        print("Cannot find FacetsTopBar")
        raise Exception("Cannot find FacetsTopBar")

    # Claims Processing > Hospital Claims Processing
    pa.press('c') # Capitation
    time.sleep(0.5)
    pa.press('c') # Claims Processing
    time.sleep(0.5)
    pa.press('right') # Claims Processing Submenu
    time.sleep(0.5)
    pa.press('h') # Hospital Claims Processing
    time.sleep(0.5)
    pa.press('enter')

    wait('openWork.png', 60) # Wait up to a minute for the Hospital Claims Processing tab to load
    if pa.locateOnScreen(os.path.join(imgLocation, 'openWork.png'), confidence=.5) is None:
        raise Exception("Hospital Claims Processing tab failed to load!")
    time.sleep(1)


def alwaysVisible():
    # Set Workspace to 'Always on Visible Workspace'
    # Prevents Facets screen from going into another workspace when screen is not responsive while pricing claim
    if pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace1.png'), confidence=.5) is not None:
        print("Found Facets Workspace.")
        rightClick('facetsWorkspace1.png')
        time.sleep(0.5)
        pa.press('down')
        time.sleep(0.5)
        pa.press('a')
        time.sleep(1)
        pa.moveTo(x = 350, y = 200)
    elif pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace2.png'), confidence=.5) is not None:
        print("Found Facets Workspace.")
        rightClick('facetsWorkspace2.png')
        time.sleep(0.5)
        pa.press('down')
        time.sleep(0.5)
        pa.press('a')
        time.sleep(1)
        pa.moveTo(x = 350, y = 200)
    elif pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace3.png'), confidence=.5) is not None:
        print("Found Facets Workspace.")
        rightClick('facetsWorkspace3.png')
        time.sleep(0.5)
        pa.press('down')
        time.sleep(0.5)
        pa.press('a')
        time.sleep(1)
        pa.moveTo(x = 350, y = 200)


def closeWarnings():
    # Check for err and or warning messages
    print("Looking for warning and error messages.")
    
    while (pa.locateOnScreen(os.path.join(imgLocation, 'warningAndErrorMessage1.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'warningAndErrorMessage2.png'), confidence=.7)):
                print("Found Err Messages and Warning messages.")
                pa.moveTo(x = 975, y = 515)
                time.sleep(1)
                pa.click(x = 1105, y = 495)
                time.sleep(1)
                pa.moveTo(x = 975, y = 100)
                time.sleep(.5)

    while (pa.locateOnScreen(os.path.join(imgLocation, 'errorMessages1.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'errorMessages1Red.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'errorMessages2.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'errorMessages2Red.png'), confidence=.7) is not None):
                print("Found Err Messages.")
                pa.moveTo(x = 955, y = 495)
                time.sleep(1)
                pa.click(x = 1110, y = 503)
                time.sleep(1)
                pa.moveTo(x = 975, y = 100)
                time.sleep(.5)

    while (pa.locateOnScreen(os.path.join(imgLocation, 'warningMessages.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'warningMessages2.png'), confidence=.7) is not None
            or
            pa.locateOnScreen(os.path.join(imgLocation, 'warningMessagesRed.png'), confidence=.7) is not None):
                print("Found Warning Messages.")
                pa.moveTo(x = 955, y = 495)
                time.sleep(1)
                pa.click(x = 1110, y = 503)
                time.sleep(1)
                pa.moveTo(x = 975, y = 100)
                time.sleep(.5)

def startScrapeF3(claim_no):
    try:
        print("Begin ""try"" in startScrapeF3")
        searchClmSuccess = searchClm(claim_no)
    except:
        print("Begin ""except"" in startScrapeF3")
        searchClmSuccess = searchClm(claim_no)

    if not searchClmSuccess:
        print(f"Searching for claim in Try - Except failed. Could not find claim {claim_no}.")
        Charges = None
        Allowed = None
        Benefit = None
        resultsToTable(claim_no, Charges, Allowed, Benefit)
        return
    time.sleep(1)

    print(f"Claim {claim_no} was found. Checking for Error pop-ups.")

    if pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle.png'), confidence=.7) is not None:
        print("Found FacetErrorTriangle.png on first check.")
        Charges = None
        Allowed = None
        Benefit = None
        pa.press('enter')
        time.sleep(1)
        if pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle.png'), confidence=.7) is not None:
            print("Found FacetErrorTriangle.png on second check.")
            pa.press('enter')
            time.sleep(1)
            if pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle.png'), confidence=.7) is not None:
                print("FacetErrorTriangle.png did not disappear.")
                raise Exception("FacetErrorTriangle did not disappear.")
        resultsToTable(claim_no, Charges, Allowed, Benefit)
        return

    if pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.7) is not None:
        print("Found facetErrorX.png on first check.")
        pa.press('enter')
        if pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.7) is not None:
            print("facetErrorX.png did not disappear.")
            raise Exception("facetErrorX did not disappear.")
        pa.press('enter')
        Charges = None
        Allowed = None
        Benefit = None
        resultsToTable(claim_no, Charges, Allowed, Benefit)
        return
    else:
        # Go to Line Items tab
        print(f"No Error pop-ups found. Going to Line Items for claim {claim_no} .")

        if pa.locateOnScreen(os.path.join(imgLocation, 'chargeAllowedBenefitLabelsOC.png'), confidence=.7) is None:
            print("Cannot find Amount Labels. Trying to open Line Items.")
            wait('lineItems.png', 60)
            pa.moveTo(x = 115, y = 185)
            time.sleep(.5)
            pa.doubleClick(x = 115, y = 185)
            time.sleep(1)
        else:
            print("Found amount labels.")

        # Check for warning message after going to Line Items
        print("Looking for warning and error messages.")
        closeWarnings()

        time.sleep(.5)
        print("Starting to Price Claim.")
        pa.press('f3')
        time.sleep(1)

        loopStartTime = time.time()

        while time.time()-loopStartTime < 900:
            if (exists('adjudicationInProcess.png')
                or
                exists('notResponding.png')):
                print("Claim is being priced. Will sleep for 5 seconds and check again.")
                time.sleep(5)         
            if (not exists('notResponding.png')
                and
                not exists('adjudicationInProcess.png')):
                # Possible that claim is still being priced, and that adjudication image may reappear.
                time.sleep(1)
                if exists('facetsBottomLeftBlank.png'):
                    print("leaving loop")
                    time.sleep(1)
                    break
                else:
                    print("Did not find facetsBottomLeftBlank.")
                    time.sleep(1)
        else:
            workspace1 = pa.center(pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace1.png'), confidence=.5))
            workspace2 = pa.center(pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace2.png'), confidence=.5))
            workspace3 = pa.center(pa.locateOnScreen(os.path.join(imgLocation, 'facetsWorkspace3.png'), confidence=.5))
                
            print("Trying to get back to Facets.")
            
            if workspace1 is not None:
                pa.click('facetsWorkspace1.png')
                time.sleep(.5)
                pa.click(workspace1)
                time.sleep(.5)
                pa.click(workspace1)
                time.sleep(.5)                    
                pa.click(workspace1)
                time.sleep(.5)
            elif workspace2 is not None:
                pa.click('facetsWorkspace2.png')
                time.sleep(.5)
                pa.click(workspace2)
                time.sleep(.5)
                pa.click(workspace2)
                time.sleep(.5)                    
                pa.click(workspace2)
                time.sleep(.5)
            elif workspace3 is not None:
                pa.click('facetsWorkspace3.png')
                time.sleep(.5)
                pa.click(workspace3)
                time.sleep(.5)
                pa.click(workspace3)
                time.sleep(.5)                    
                pa.click(workspace3)
                time.sleep(.5)
            else:
                raise Exception("Repricing took longer than 15 minutes!")

        print("Exited out of startScrape While Loop")

        print("Looking for warning messages after pricing.")
        closeWarnings()

        print("Looking for amountAllowed .")
        
        wait('chargeAllowedBenefitLabelsOC.png', 60)
        if pa.locateOnScreen(os.path.join(imgLocation, 'chargeAllowedBenefitLabelsOC.png'), confidence=.7) is not None:
            tempAllowed = Region(x = 410, y = 948, w = 130, h = 17)
            dateconfig = {"mode": "block", "character_whitelist": "1234567890$"}
            amountAllowed = tempAllowed.text(dateconfig).replace("$", "")
            print(f"Found Amount Allowed {amountAllowed} for claim {claim_no} .")
        else:
            raise Exception(f"Cannot find Allowed benefit for claim {claim_no}!")

        if len(amountAllowed) < 3:
            Charges = None
            Allowed = None
            Benefit = None
            resultsToTable(claim_no, Charges, Allowed, Benefit)
        else:
            try:
                amountFloat = float(amountAllowed)
                amountDec = amountFloat * 0.01
                amountAllowed = "{:.2f}".format(amountDec)
                Charges = None
                Allowed = amountAllowed
                Benefit = None
                resultsToTable(claim_no, Charges, Allowed, Benefit)
            except:
                Charges = None
                Allowed = None
                Benefit = None
                resultsToTable(claim_no, Charges, Allowed, Benefit)

        print(f"Finished startScrape for claim {claim_no} .")


def searchClm(claim_no):
    print(f"Searching for {claim_no} .")
    pa.moveTo(x = 430, y = 117)
    time.sleep(.5)
    pa.click(x = 430, y = 117) #Hospital Claims Processing tab
    time.sleep(.5)
    pa.hotkey('ctrl', 'o')
    time.sleep(1)
    loopStarttime = time.time()
    found = False

    while not exists('openClaimID_OC.png'):
        print("Starting while loop in searchClm.")
        if exists('openClaimID_OC.png'):
            print("Found openClaimID_OC image. Exiting loop.")
            found = True
            break
        if not exists('openClaimID_OC.png'):
            print("Cant find openClaim image. Trying to open claim search.")
            pa.moveTo(x = 430, y = 117)
            time.sleep(.5)
            pa.click(x = 430, y = 117) #Hospital Claims Processing tab
            time.sleep(.5)
            pa.hotkey('ctrl', 'o')
            time.sleep(1)
        if time.time()-loopStarttime > 180:
            raise Exception(f"While Loop reached timeout! Could not find claim search window for claim {claim_no}!")
        if pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.7) is not None:
            print("Found facetErrorX image. Closing pop up and trying search again.")
            pa.press('enter')
            time.sleep(1)
            pa.click(x = 1690, y = 230)
            time.sleep(0.5)
            pa.hotkey('ctrl', 'o')
            time.sleep(0.5)

    if not found:
        wait('openClaimID_OC.png', 10)
        if pa.locateOnScreen(os.path.join(imgLocation, 'openClaimID_OC.png'), confidence=.5) is not None:
            found = True
        else:
            raise Exception(f"Could not find claim search window for claim {claim_no}!")

    time.sleep(0.5)
    type(claim_no)
    time.sleep(0.5)
    pa.press('enter')
    Region(700, 400, x2 = 1250,y2 = 700).waitVanish('openClaimID_OC.png', 60)
    time.sleep(1)
    print(f"Checking for warning messages before pricing claim {claim_no}.")

    if (pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle.png'), confidence=.5) is not None
        or
        pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.5) is not None):
            print(f"Found Facet Error image. Checking to see if claim {claim_no} is Read-Only.")
            time.sleep(0.5)
            readonly = Region(700, 400, x2 = 1250,y2 = 700).text()
            if readonly != 'Read-Only?':
                print(f"Claim {claim_no} is being modified by another user.")
                pa.press('enter')
                time.sleep(1)
                return False
            if pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.5) is not None:
                raise Exception("Claim Search Error box did not disappear!")
            time.sleep(1)
            pa.press('enter')
            if (pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle.png'), confidence=.5) is not None
                or
                pa.locateOnScreen(os.path.join(imgLocation, 'facetErrorX.png'), confidence=.5) is not None):
                raise Exception("Claim Search Restrictions box did not disappear!")
    
    if (pa.locateOnScreen(os.path.join(imgLocation, 'openFailed.png'), confidence=.5) is not None
        or
        pa.locateOnScreen(os.path.join(imgLocation, 'facetsTriangle.png'), confidence=.5) is not None):
            pa.press('enter')
            time.sleep(1)
            pa.moveTo(x = 960, y = 485)
            time.sleep(.5)
            pa.click(x = 1085, y = 600)
            time.sleep(1)
            print(f"Claim {claim_no} is being modified by another user.")
            return False

    if pa.locateOnScreen(os.path.join(imgLocation, 'fileReservation.png'), confidence=.5) is not None:
        pa.press('enter')
        print(f"Claim {claim_no} is being modified by another user.")
        return False

    if pa.locateOnScreen(os.path.join(imgLocation, 'FacetErrorTriangle2.png'), confidence=.5) is not None:
        pa.press('enter')
        print(f"Open Failed error message for Claim {claim_no} .")
        return False

    if (pa.locateOnScreen(os.path.join(imgLocation, 'hipaaPrivacyPopup.png'), confidence=.5) is not None
        or
        pa.locateOnScreen(os.path.join(imgLocation, 'HIPAA.png'), confidence=.5) is not None):
            pa.press('enter')

    closeWarnings()

    wait('mainCLMButtonsOC.png', 60)
    print("Found mainCLMButtons .")
    time.sleep(0.5)
    return True

def resultsToTable(claim_no, Charges, Allowed, Benefit):
    Date_Insert = datetime.now()
    connection = pymssql.connect(host = 'WP000075696.ms.ds.uhc.com', database = 'RacerResearch', user = msid, password = password)
    cursor = connection.cursor()
    query = "INSERT INTO racerresearch.[DBDataAnalytics-DM].OptumCare_Facets_F3_Pricer_Results (CLAIM_NO, Charges, Allowed, Benefit, Date_Insert, MSID) values(%s,%s,%s,%s,%s,%s)"
    cursor.execute(query, (claim_no, Charges, Allowed, Benefit, Date_Insert, UserMSID))
    connection.commit()
    connection.close()

#endregion FUNCTIONS

#region MAIN

if __name__ == "__main__":
    try:
        setupFacets()
        print("Finished setting up facets")
        print("Getting inventory")
        queueInventory = scrapeQ(msid, password, 'WP000075696.ms.ds.uhc.com', projectid, scrapeQName, mainuuid=scriptuuid, userMSid=UserMSID) #, userMSid=UserMSID
        scrapeStartTime = datetime.now()

        while queueInventory.nextItem():
            body = queueInventory.getItem()
            claim_no = body
            print(f"Retrieved claim {claim_no} from the queue. Going to startScrapeF3.")
            time.sleep(1)
            startScrapeF3(claim_no)
        print("No more claims in the queue.")

        scrapeEndTime = (datetime.now() - scrapeStartTime)
        print(f"Scrape ran for {scrapeEndTime} .")
        time.sleep(1)
        pa.hotkey('alt', 'f')
        time.sleep(0.5)
        pa.press('up')
        pa.press('enter') 
    except:
        errorTime = datetime.now().strftime('%m_%d_Y %H_%M_%S')
        with open(f'/home/headless/errorLog {errorTime}.log', 'w') as f:
            f.write(traceback.format_exc())
        print(traceback.format_exc())
        exit(1)

#endregion MAIN
