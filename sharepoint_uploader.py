# Script by: DarkSplash
# Last edited: 2023-01-12

# This is a 100% working Python script that will upload a local file to a 
# SharePoint/OneDrive/Teams location. This script supports MFA & non-MFA logins,
# and requires quite alot of setup before it will work. Please reference the 
# README.md file for all script setup, as most of the comments inside this 
# script try to explain the more technical side of what is happening.

import core.dotenv_checker as dotenv_checker            # Script to check msal_config.env variables
import core.token_generator as token_generator          # Script to generate a MSAL token
import core.driveid_finder as driveid_finder            # Script to attempt to find a SharePoint/OneDrive/Teams drive_id

import argparse
import os
import requests
import urllib

from pathlib import Path

# Packageless Terminal Colors: https://stackoverflow.com/a/21786287
RED = "\x1b[1;31;40m"
GREEN = "\x1b[1;32;40m"
BLUE = "\x1b[1;34;40m"
CLEAR = "\x1b[0m"



def argparseInit():
    """
    Function for command line flags that can be added while running the script.
    Currently, only flag is -G or --gui, which will launch Firefox with a GUI
    instead of it launching headlessly how it normally does.

    Returns
    -------
    guiFlag : bool
        A flag that will be set to True if the user wishes to run the selenium
        portion of this script with a GUI instead of headlessly. By default set
        to False, and can be set to True with the -G or --gui args.
    useMFA : bool
        A flag that will be set to False if the user wishes to run the script
        with an account with no MFA setup. By default set to True, and can be
        set to False with the -N or --nomfa args.
    rundDriveID : bool
        A flag that will be set to True if the user wishes to attempt to find
        their drive id through the script. By default set to False, and can be
        set to True with the -D or --driveid args.
    """
    guiFlag = False
    useMFA = True
    runDriveID = False

    parser = argparse.ArgumentParser()
    parser.add_argument("-G","--gui", help="Runs the Selenium/Firefox portion of this script with a GUI instead of headlessly", action="store_true")
    parser.add_argument("-N","--nomfa", help="Allows you to run the script without filling in the MFA_SECRET variable", action="store_true")
    parser.add_argument("-D","--driveid", help="Runs two different methods to attempt to find your M365_DRIVE_ID variable", action="store_true")
    args = parser.parse_args()

    if args.gui:
        print("\nFirefox will launch with a GUI instead of headlessly...")
        guiFlag = True
    if args.nomfa:
        print("\nScript will not check for MFA...")
        useMFA = False
    if args.driveid:
        print("\nScript will only attempt to generate drive_id's...")
        runDriveID = True
    if guiFlag == False and useMFA == True and runDriveID == False:
        print(f"\n{BLUE}Optional runtime argument can be displayed by adding the \'-h\' flag to the end of your python command above.{CLEAR}")
    
    return guiFlag, useMFA, runDriveID



def uploadFile(token):
    headers = {'Authorization': 'Bearer {}'.format(token['access_token'])}  # Header will be used for authentication with Microsoft Graph

    fullRelativePath = urllib.parse.quote(f'{os.environ.get("M365_FOLDER_PATH")}/{os.environ.get("M365_FILENAME")}')
    fileRelativePath = urllib.parse.quote(f'{os.environ.get("M365_FILENAME")}')
    folderRelativePath = urllib.parse.quote(f'{os.environ.get("M365_FOLDER_PATH")}')

    # Checking to see if file exists
    result = requests.get(f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/root:/{fullRelativePath}', headers=headers)
    if result.status_code == 200:
        fileExists = True
        fileID = result.json()['id']
    else:
        fileExists = False
        fileID = ''

    # Getting folder ID
    result = requests.get(f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/root:/{folderRelativePath}', headers=headers)
    folderID = result.json()['id']

    # Getting local filesize
    stringPath = f'{os.getcwd()}/{os.environ.get("M365_FILENAME")}'

    if Path(stringPath).exists():
        uploadPath = Path(stringPath)
        st = os.stat(uploadPath)
        size = st.st_size

    if size <= 4194304:
        if fileExists:
            result = requests.put(
            f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/items/{fileID}/content',
            headers = headers,
            data=open(os.environ.get("M365_FILENAME"), 'rb').read()
            )
        else:
            result = requests.put(f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/items/{folderID}:/{fileRelativePath}:/content'
                            ,headers = headers
                            ,data = open(os.environ.get("M365_FILENAME"), 'rb').read()
                                )
    else:
        result = requests.post(
        f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/items/{folderID}:/{fileRelativePath}:/createUploadSession',
        headers={'Authorization': 'Bearer ' + token['access_token']},
        json={
            '@microsoft.graph.conflictBehavior': 'replace',
            'description': 'Uploading a large file',
            'fileSystemInfo': {'@odata.type': 'microsoft.graph.fileSystemInfo'},
            'name': os.environ.get("M365_FILENAME")
            }
            )
        upload_url = result.json()['uploadUrl']
        CHUNK_SIZE = 10485760
        chunks = int(size / CHUNK_SIZE) + 1 if size % CHUNK_SIZE > 0 else 0
        with open(os.environ.get("M365_FILENAME"), 'rb') as fd:
            start = 0
            for chunk_num in range(chunks):
                chunk = fd.read(CHUNK_SIZE)
                bytes_read = len(chunk)
                upload_range = f'bytes {start}-{start + bytes_read - 1}/{size}'
                result = requests.put(upload_url,
                        headers={
                            'Content-Length': str(bytes_read),
                            'Content-Range': upload_range
                        },
                        data=chunk
                    )
                result.raise_for_status()
                start += bytes_read
    
    fileCheck = requests.get(f'https://graph.microsoft.com/v1.0/drives/{os.environ.get("M365_DRIVE_ID")}/root:/{fullRelativePath}', headers=headers)
    if fileCheck.status_code == 200:
        print(f"\n{GREEN}File \"{os.environ.get('M365_FILENAME')}\" has been sucessfully uploaded!{CLEAR}")
    else:
        print(f"\n{RED}File \"{os.environ.get('M365_FILENAME')}\" has not been sucessfully uploaded!{CLEAR}")
       


def main():
    guiFlag, useMFA, runDriveID = argparseInit()        # Checking for command flags
    dotenv_checker.dotenvInit(useMFA, runDriveID)

    token = token_generator.tokenGen(guiFlag, useMFA)

    if runDriveID:                                      # If the flag has been set to programatically check for drive_id's
        driveid_finder.findDriveID(token)
        raise SystemExit(0)                             # Exiting the script as none of the variables needed to download the file were checked

    print("\nUploading file...")
    uploadFile(token)                                   # Upload the file using the token for authentication



if __name__ == "__main__":
    main()