# Script by: DarkSplash
# Last edited: 2023-01-03

# This is one of the dependency scripts for the downloader/uploader script.
# This specific script checks to make sure you have either properly configured
# your msal_config.env file, or will ask if you want the script to make you a
# template msal_config.env file for you to use during setup.

from dotenv import load_dotenv
from pathlib import Path

import os

# Packageless Terminal Colors: https://stackoverflow.com/a/21786287
RED = "\x1b[1;31;40m"
GREEN = "\x1b[1;32;40m"
CLEAR = "\x1b[0m"



def msalConfigChecker(useMFA: bool, runDriveID: bool):
    """
    Function to check all variables in the msal_config.env file to make sure they
    are not blank, and if it is a variable that should be a certain length, checks
    to make sure that variable is at the proper length.

    Parameters
    ----------
    useMFA : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if MFA script procedures will be ran. By default this is set to True.
        If set to False, the MFA_SECRET variable is ignored in this function.
    runDriveID : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if the script will only attempt to find drive_id's. By default this is
        set to False. If set to True, this function will only check the 
        msal_config.env variables needed to generate a token.

    Returns
    -------
    emptyVars : bool
        A flag that will be set to true if any of the variables are
        misconfigured. Used in dotenvInit() in an if statement to check 
        msal_config.env's file state.
    """
    emptyVars = False

    if runDriveID:
        try:
            if len(os.environ.get("CLIENT_ID")) == 0 or len(os.environ.get("CLIENT_ID")) != 36:
                print(f"\n{RED}CLIENT_ID{CLEAR} variable empty or improper length")
                print(f"CLIENT_ID Length: {RED}{len(os.environ.get('CLIENT_ID'))}{CLEAR}")
                print(f"Proper Length: {GREEN}36{CLEAR}")
                emptyVars = True
        except:
            print(f"\n{RED}CLIENT_ID{CLEAR} variable missing from msal_config.env")
            emptyVars = True
        try:
            if len(os.environ.get("AUTHORITY_URL")) == 0 or len(os.environ.get("AUTHORITY_URL")) != 70:
                print(f"\n{RED}AUTHORITY_URL{CLEAR} variable empty or improper length")
                print(f"AUTHORITY_URL Length: {RED}{len(os.environ.get('AUTHORITY_URL'))}{CLEAR}")
                print(f"Proper Length: {GREEN}70{CLEAR}")
                emptyVars = True
        except:
            print(f"\n{RED}AUTHORITY_URL{CLEAR} variable missing from msal_config.env")
            emptyVars = True
        try:
            if len(os.environ.get("MFA_SECRET")) == 0 and useMFA:
                print(f"\n{RED}MFA_SECRET{CLEAR} variable empty")
                print(f"If you wish to run the script without MFA, add the flag {GREEN}-N{CLEAR} or {GREEN}--nomfa{CLEAR} to the end of")
                print("your python command. Look at the beginning of the README for more details and an example.")
                emptyVars = True
        except:
            if useMFA:                                  # If MFA should still be used, report the missing variable
                print(f"\n{RED}MFA_SECRET{CLEAR} variable missing from msal_config.env")
                emptyVars = True
        try:
            if len(os.environ.get("M365_USERNAME")) == 0:
                print(f"\n{RED}M365_USERNAME{CLEAR} variable empty")
                emptyVars = True
        except:
            print(f"\n{RED}M365_USERNAME{CLEAR} variable missing from msal_config.env")
            emptyVars = True
        try:
            if len(os.environ.get("M365_PASSWORD")) == 0:
                print(f"\n{RED}M365_PASSWORD{CLEAR} variable empty")
                emptyVars = True
        except:
            print(f"\n{RED}M365_PASSWORD{CLEAR} variable missing from msal_config.env")
            emptyVars = True
        
        return emptyVars

    try:
        if len(os.environ.get("CLIENT_ID")) == 0 or len(os.environ.get("CLIENT_ID")) != 36:
            print(f"\n{RED}CLIENT_ID{CLEAR} variable empty or improper length")
            print(f"CLIENT_ID Length: {RED}{len(os.environ.get('CLIENT_ID'))}{CLEAR}")
            print(f"Proper Length: {GREEN}36{CLEAR}")
            emptyVars = True
    except:
        print(f"\n{RED}CLIENT_ID{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("AUTHORITY_URL")) == 0 or len(os.environ.get("AUTHORITY_URL")) != 70:
            print(f"\n{RED}AUTHORITY_URL{CLEAR} variable empty or improper length")
            print(f"AUTHORITY_URL Length: {RED}{len(os.environ.get('AUTHORITY_URL'))}{CLEAR}")
            print(f"Proper Length: {GREEN}70{CLEAR}")
            emptyVars = True
    except:
        print(f"\n{RED}AUTHORITY_URL{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("M365_DRIVE_ID")) == 0 or len(os.environ.get("M365_DRIVE_ID")) != 66:
            print(f"\n{RED}M365_DRIVE_ID{CLEAR} variable empty or improper length")
            print(f"M365_DRIVE_ID Length: {RED}{len(os.environ.get('M365_DRIVE_ID'))}{CLEAR}")
            print(f"Proper Length: {GREEN}66{CLEAR}")
            emptyVars = True
    except:
        print(f"\n{RED}M365_DRIVE_ID{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("M365_ITEM_PATH")) == 0:
            print(f"\n{RED}M365_ITEM_PATH{CLEAR} variable empty")
            emptyVars = True
    except:
        print(f"\n{RED}M365_ITEM_PATH{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("MFA_SECRET")) == 0 and useMFA:
            print(f"\n{RED}MFA_SECRET{CLEAR} variable empty")
            print(f"If you wish to run the script without MFA, add the flag {GREEN}-N{CLEAR} or {GREEN}--nomfa{CLEAR} to the end of")
            print("your python command. Look at the beginning of the README for more details and an example.")
            emptyVars = True
    except:
        if useMFA:                                      # If MFA should still be used, report the missing variable
            print(f"\n{RED}MFA_SECRET{CLEAR} variable missing from msal_config.env")
            emptyVars = True
    try:
        if len(os.environ.get("M365_USERNAME")) == 0:
            print(f"\n{RED}M365_USERNAME{CLEAR} variable empty")
            emptyVars = True
    except:
        print(f"\n{RED}M365_USERNAME{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("M365_PASSWORD")) == 0:
            print(f"\n{RED}M365_PASSWORD{CLEAR} variable empty")
            emptyVars = True
    except:
        print(f"\n{RED}M365_PASSWORD{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    try:
        if len(os.environ.get("M365_FILENAME")) == 0:
            print(f"\n{RED}M365_FILENAME{CLEAR} variable empty")
            emptyVars = True
    except:
        print(f"\n{RED}M365_FILENAME{CLEAR} variable missing from msal_config.env")
        emptyVars = True
    
    return emptyVars



def msalConfigCreator():
    """
    Function creates a blank msal_config.env file with all the necessary
    variables already added in and properly spaced.
    """
    with open("msal_config.env", "w") as file:
        file.write("CLIENT_ID =         \"\"\n")
        file.write("AUTHORITY_URL =     \"\"\n")
        file.write("M365_DRIVE_ID =     \"\"\n")
        file.write("M365_ITEM_PATH =    \"\"\n")
        file.write("M365_FILENAME =     \"\"\n")
        file.write("MFA_SECRET =        \"\"\n")
        file.write("M365_USERNAME =     \"\"\n")
        file.write("M365_PASSWORD =     \"\"\n")



def dotenvInit(useMFA: bool, runDriveID: bool):
    """
    Function loads the msal_config.env file that should be created during the
    setup process. If it does not exist, it will ask the user if they want to
    create a blank template msal_config.env file and exits the script after the
    user's response. If the msal_config.env file does exist, it will check to
    make sure that all of the variables are filled out. If a variable is blank
    or the wrong length, it will report the variable and exit the script.

    Parameters
    ----------
    useMFA : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if MFA script procedures will be ran. By default this is set to True.
        Passed to msalConfigChecker() in this function.
    runDriveID : bool
        A boolean variable that is set at script runtime with a flag. Determines
        if the script will only attempt to find drive_id's. By default this is
        set to False. Passed to msalConfigChecker() in this function.

    Raises
    ------
    SystemExit
        Exits the script if either msal_config.env doesn't exist or the
        variables within the file are blank or misconfigured.
    """
    print("\nChecking msal_config.env file setup...")
    stringPath = f"{os.getcwd()}/msal_config.env"
    dotenvPath = Path(stringPath)                       # Converting into proper path for whatever OS the script is on
    
    if Path(stringPath).exists():                       # If the file exists
        load_dotenv(dotenvPath)                         # Loading the environment variables

        if msalConfigChecker(useMFA, runDriveID):       # Checking to see if vars are populated
            print("\nOne or more of your variables in msal_config.env is empty, misconfigured, or missing.")
            print("Add/fix the data listed above in red and run the script again.\n")
            raise SystemExit(0)
        else:
            print(f"{GREEN}Your msal_config.env file appears to be properly configured!{CLEAR}\n")
    else:
        print("You do not have msal_config.env in the same directory as your script")
        print("or you are executing the Python code from a different directory than")
        print("the one that has your msal_config.env file (if it has already been made).")
        print(f"\nDirectory the script is checking: {RED}{dotenvPath}{CLEAR}\n")
        while True:
            flag = input("Do you want a blank template msal_config.env file to be made (yes/no)?\n")
            
            if "yes" in flag.lower():                   # Creating blank msal_config.env
                print("\nCreating blank msal_config.env file...")
                msalConfigCreator()
                print("File created! Please now follow the README instructions on filling out the variables.")
                print("Exiting script...")
                raise SystemExit(0)
            elif "no" in flag.lower():                  # Letting user create msal_config.env
                print("\nPlease make a msal_config.env file following the README instructions.")
                print("Exiting script...")
                raise SystemExit(0)
            else:
                print("\nPlease either answer \"yes\" or \"no\".\n\n")
                continue