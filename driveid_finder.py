# Script by: DarkSplash
# Last edited: 2023-01-03

# This is one of the dependency scripts for the downloader/uploader script.
# This specific script has only one purpose, to attempt to find your drive_id
# variable via two methods. When you select this flag at runtime, certain
# msal_config.env variables don't get checked for, and after generating your token,
# only this script will run, spit out the two attempts, and will exit the script.

import requests



def findDriveID(token):
    """
    Function prints out two different Microsoft Graph API's that may provide
    you your M365_DRIVE_ID variable.  Will only run if the -D or --driveid
    flags/args are added at script runtime.

    Parameters
    ----------
    token : dict
        A dictionary object created by MSAL's acquire_token_by_auth_code_flow()
        function. Contains information needed to create the HTTP header that is
        used for authentication with the Microsoft Graph API calls.
    """
    headers = {'Authorization': 'Bearer {}'.format(token['access_token'])}  # Header will be used for authentication with Microsoft Graph

    result = requests.get(f'https://graph.microsoft.com/v1.0/drive/microsoft.graph.recent()', headers=headers)  # Attempt for drive_id by looking at recent files
    result2 = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe', headers=headers)          # Attempt for drive_id by looking at files shared with account
    resultJSON = result.json()
    resultJSON2 = result2.json()

    print("\nRecently Viewed Files:")
    print(resultJSON)
    print("\nFiles \"Shared With Me\":")
    print(resultJSON2)
    print("\n\nThe two JSON outputs above may or may not have your driveId variable somewhere in there.")
    print("I would recommend copy-pasting the outputs into some online JSON formatter so they are easier to read.")