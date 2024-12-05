from dotenv import load_dotenv
import pandas as pd
from slack_sdk import WebClient
import os

import slack_sdk

load_dotenv()
client = WebClient(token=os.getenv('SLACK_USER_TOKEN'))  

df = pd.read_excel('books2.xlsx', sheet_name='behind')

mails = df['mail'].tolist()
names = df['name'].tolist()


user_ids = []


for mail in mails:
    try:
        user_info = client.users_lookupByEmail(email=mail)
        if user_info['ok']:
            user_id = user_info['user']['id']
            user_ids.append(user_id)
        else:
            user_ids.append(None)
            print(f"Error searching for {mail}: {user_info['error']}")
    except slack_sdk.errors.SlackApiError as e:
        print(f"Error searching for {mail}: {e.response['error']}")
        user_ids.append(None)



for name, mail, user_id in zip(names, mails, user_ids):
    if user_id:
        message = f""" I encourage you to book an appointment with me so we can meet and go through all the topics that are causing a roadblock on your path.

        """
        response = client.chat_postMessage(channel=user_id, text=f" {message}")
        if response['ok']:
            print(f"Message sent to {name} ({mail})")
        else:
            print(f"Error with {name} ({mail}):", response['error'])
    else:
        print(f"User ID not found for {mail}.")
