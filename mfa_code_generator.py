# Script by: DarkSplash
# Last edited: 2023-01-12

# Small script to print out your six-digit MFA code to console so you don't have
# to setup MFA on your application of choice.  Only requires the "MFA_SECRET"
# variable to be filled out in msal_config.env.

from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path

import os
import pyotp


def getTOTP(mfa_secret: str) -> str:
    """
    Function takes a MFA secret and generates you the Time-Based One Time 
    Password (TOTP) using the PyOTP library. Returns a six-digit TOTP code.

    Parameters
    ----------
    mfa_secret : str
        The secret key that gets generated when setting up MFA. Note that this
        string is from the general Authenticator App when looking at your O365
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



def dotenvInit():
    """
    Function loads the msal_config.env file that should created during the
    setup process. If it does not exist, it directs the user to the
    sharepoint_downloader_msal.py file. If it does exist, it only checks
    for the MFA_SECRET variable as that is the only thing needed to
    generate a TOTP code.
    """
    stringPath = f"{os.getcwd()}/msal_config.env"
    if Path(stringPath).exists():                       # If the file exists
        dotenvPath = Path(stringPath)                   # Converting into proper path for whatever OS the script is on
        load_dotenv(dotenvPath)                         # Loading the environment variables

        if len(os.environ.get("MFA_SECRET")) == 0:      # If the secret variable is empty exit the script
            print("MFA Secret variable is empty.  Please enter your MFA secret in msal_config.env and try again.")
            raise SystemExit(0)
    else:
        print("You do not have msal_config.env in the same directory as your script.")
        print("Please either create it following the steps in the README or run")
        print("sharepoint_downloader_msal.py to create a blank msal_config.env file.")
        raise SystemExit(0)



def main():
    dotenvInit()
    now = datetime.now()
    
    print(f"{getTOTP(os.environ.get('MFA_SECRET'))}")
    print(f"Code valid for {30 - now.second%30} more seconds")  # Timer on TOTP is 30 seconds long, resetting at 30 sec and next minute



if __name__ == "__main__":
    main()