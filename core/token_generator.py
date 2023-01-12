# Script by: DarkSplash
# Last edited: 2023-01-06

# This is one of the dependency scripts for the downloader/uploader script.
# This specific script first checks to see if all selenium dependencies are installed,
# and then later logs into your M365 account to generate the MSAL token.

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options as FirefoxOptions

import msal
import os
import pyotp
import time

# Packageless Terminal Colors: https://stackoverflow.com/a/21786287
RED = "\x1b[1;31;40m"
GREEN = "\x1b[1;32;40m"
CLEAR = "\x1b[0m"



def seleniumChecker():
    """
    Function to quickly check if Selenium, Firefox, and geckodriver are all
    setup properly. If they aren't, the script instructs the user on the likely
    errors and points them in the direction of the README.

    Raises
    ------
    SystemExit
        Exits the script if Selenium does not launch.
    """
    print("Checking selenium configuration...")
    try:
        ffOpt = FirefoxOptions()
        ffOpt.add_argument("-headless")
        driver = webdriver.Firefox(options=ffOpt)
        driver.quit()

        print(f"{GREEN}Selenium and Firefox/geckodriver appear to be properly configured!{CLEAR}")
    except Exception as e:
        print(f"\nSelenium Error:\n{e}")
        
        if "executable needs to be in PATH" in str(e):  # Checking to see if it is the error when missing geckodriver file
            print(f"{RED}Very high likelihood you are missing geckodriver in the script's directory{CLEAR}")
            print(f"Add/check the geckodriver file in {RED}{os.getcwd()}{CLEAR}")
            print("or execute the code in a directory that already has the geckodriver file in it.")
        
        print("\nAn error occurred with Selenium. Make sure you have both Firefox installed and")
        print("a geckodriver executable that both matches your Firefox version and also is in")
        print("the same directory as this script or part of your PATH/Environment Variable.")
        print("Look at the README for further instructions on how to do this.")
        print("\nIf this is your first time running the script, try running it once more.")
        print("\nExiting script...")
        raise SystemExit(0)



def getTOTP(mfa_secret: str) -> str:
    """
    Function takes an MFA secret and generates you the Time-Based One Time
    Password (TOTP) using the PyOTP library. Returns a six-digit TOTP code.

    Parameters
    ----------
    mfa_secret : str
        The secret key that gets generated when setting up MFA. Note that this
        string is from the general Authenticator App when looking at your M365
        account security tab, not the Microsoft Authenticator. You can have 
        both a Microsoft Authenticator and a general Authenticator App setup at
        the same time.

    Returns
    -------
    value : str
        The six-digit TOTP value that correlates to the current time and the
        mfa_secret.  Each value is valid for 30 seconds, with values resetting
        at the start of each minute and halfway through each minute.
    """
    totp = pyotp.TOTP(mfa_secret)
    value = totp.now()
    return value



def loginProcess(authFlow: dict, guiFlag: bool, useMFA: bool) -> str:
    """
    Function takes an auth flow generated by MSAL's initiate_auth_code_flow()
    function and uses Selenium to login and accept the Azure app permissions.
    After logging in, the redirect URL will be returned for future use.

    This is by far the jankiest function in this whole script, and could break
    on a whim due to Google updating/changing their login UI, so if in the future
    this script doesn't work, this is my prime suspect as the culprit.

    Parameters
    ----------
    authFlow : dict
        A dictionary object generated by MSAL's initiate_auth_code_flow().
        Contains information about what authentication URL to use and other 
        internal information needed to eventually generate a token.
    guiFlag : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if Firefox will open with a GUI or not. By default this is set to False.
    useMFA : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if MFA script procedures will be ran. By default this is set to True.

    Returns
    -------
    url : str
        This URL contains the Azure App redirect URI plus the auth_response dict
        in string format. Since the App Registration should redirect to a bogus
        localhost address, the script grabs that URL and returns it to be
        converted into a dict in the createAuthResponseDict() function.
    """
    sleepDuration = 5                                   # How long the script waits for new elements to actually load on the site
    driverWaitDuration = 3                              # How long the webdriver will look for the new element
    
    ffOpt = FirefoxOptions()
    ffOpt.add_argument("-headless")                     # Option for Firefox without a GUI

    if guiFlag:
        driver = webdriver.Firefox()                    # Launching Firefox with a GUI
    else:
        driver = webdriver.Firefox(options=ffOpt)       # Launching Firefox without a GUI
    
    driver.get(authFlow["auth_uri"])                    # Opening up authentication page

    ### PASSWORD INPUT FIELD ***
    time.sleep(sleepDuration)                           # Waiting for the webpage to load
    passwordBox = WebDriverWait(driver, driverWaitDuration).until(               
        EC.presence_of_element_located((By.ID, "i0118")))   # Selecting password input box
    
    passwordBox.send_keys(os.environ.get("M365_PASSWORD"))
    passwordBox.send_keys(Keys.RETURN)

    ### MFA CODE FIELD ###
    if useMFA:
        time.sleep(sleepDuration)                       # Waiting for redirect to MFA auth page
        try:                                            # Sometimes MFA doesnt appear even with MFA enabled?
            otpBox = WebDriverWait(driver, driverWaitDuration).until(
                EC.presence_of_element_located((By.ID, "idTxtBx_SAOTCC_OTC")))  # Selecting MFA input box

            mfaCode = getTOTP(os.environ.get("MFA_SECRET")) # Grab TOTP code only after page has loaded due to time sensitive nature of TOTPs
            otpBox.send_keys(mfaCode)
            otpBox.send_keys(Keys.RETURN)
        except:                                         # If for whatever reason the MFA login doesn't appear, pass on by
            pass

    ### APP PERMISSIONS FIELD ###
    time.sleep(sleepDuration)                           # Accepting Azure App permissions
    try:
        firstTimeCheck = WebDriverWait(driver, driverWaitDuration).until(
            EC.presence_of_element_located((By.ID, "loginHeader")))

        if firstTimeCheck.text == "Permissions requested":  # If this is the first time running after AD app creation, accept app permissions
            acceptButton = WebDriverWait(driver, driverWaitDuration).until(
                EC.presence_of_element_located((By.ID, "idSIButton9")))
            acceptButton.click()
    except:
        pass

    ### REMEMBER THIS PC ###
    time.sleep(sleepDuration)
    try:                                                # If it goes to the Remember this PC prompt (unable to check if this actually works)
        noButton = WebDriverWait(driver, driverWaitDuration).until( # Cannot get this prompt to re-appear no matter what I try
            EC.presence_of_element_located((By.ID, "idBtn_Back")))
        noButton.click()
    except:                                             # If there is no "Remember this PC" prompt, continue on to the Redirect URI
        pass

    time.sleep(10)                                      # Waiting for Firefox to fail loading the localhost redirect
    url = driver.current_url                            # Grabbing the URL after the redirect fails
    driver.quit()                                       # Exiting the browser

    return url                                          # Returning URL which has dict in string format



def createAuthResponseDict(url: str) -> dict:
    """
    Function takes the url obtained from loginProcess() and converts it into a
    dict by doing numerous substring calculations.

    Parameters
    ----------
    url : str
        String returned by the loginProcess() function. Contains all the data
        in an annoying string format.

    Returns
    -------
    authResponse : dict
        Dictionary that gets created from the URL string. Will be used in
        MSAL's acquire_token_by_auth_code_flow() to generate a token.
        The dictionary's structure looks like the following:\n\n
        authResponse = {
            "code" :            "...substring...",
            "client_info" :     "...substring...",
            "state" :           "...substring...",
            "session_state" :   "...substring..."
        }
    """
    authResponse = {}                                   # Dict that would be created if this didn't point to a non-existent localhost webserver

    codeStart = url.find("?code=") + len("?code=")      # Making so substring grabs after the equals sign
    codeEnd = url.find("&client_info=")                 # Last index is where the next variable starts
    authResponse["code"] = url[codeStart:codeEnd]       # Add substring to dict, this repeats three more times below...

    infoStart = url.find("&client_info=") + len("&client_info=")
    infoEnd = url.find("&state=")
    authResponse["client_info"] = url[infoStart:infoEnd]

    stateStart = url.find("&state=") + len("&state=")
    stateEnd = url.find("&session_state=")
    authResponse["state"] = url[stateStart:stateEnd]

    sessionStart = url.find("&session_state=") + len("&session_state=")
    tmpStr = url[sessionStart:]                         # Taking precautions in case octothorpes appear anywhere else in the string
    sessionEnd = tmpStr.find("#")
    authResponse["session_state"] = tmpStr[:sessionEnd]

    return authResponse


def tokenGen(guiFlag: bool, useMFA: bool):
    appScopes = ["Files.ReadWrite.All","Sites.Read.All"]        # Scopes defined in Azure App Registration
    seleniumChecker()                                           # Making sure Selenium & Firefox/geckodriver work
    
    pca = msal.PublicClientApplication(os.environ.get("CLIENT_ID"), authority=os.environ.get("AUTHORITY_URL"))  # Create a Public Application
    authFlow = pca.initiate_auth_code_flow(appScopes, login_hint=os.environ.get("M365_USERNAME"))               # Generate the auth flow

    print("\nLogging into M365 and accepting app permissions (can take up to a minute)...")
    authResponseUrl = loginProcess(authFlow, guiFlag, useMFA)   # Get the auth response in string format
    authResponse = createAuthResponseDict(authResponseUrl)      # Convert the auth response string into a dict

    print("\nGenerating token..")
    token = pca.acquire_token_by_auth_code_flow(auth_code_flow=authFlow, auth_response=authResponse)    # Generate a token with the authFlow and authResponse dictionaries

    return token