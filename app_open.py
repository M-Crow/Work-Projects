# script to open windows apps on login
import os
import webbrowser
import datetime as dt
import win32com.client

# opens outlook
os.startfile("Outlook")
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
root_folder = mapi.Folders['Support'].Folders['Inbox']

# List the current time in 12 hr string format and replacing ":" and "/" with "_"
current_time = str(dt.datetime.now().strftime("%m/%d/%Y-%I:%M %p"))
current_time = current_time.replace(':', '_').replace('/', '_')

# this is set to the current time
date_time = dt.datetime.now()
#This is set to 24 hours ago; you can change timedelta's argument to whatever you want it to be
delta_date_time = dt.datetime.now() - dt.timedelta(hours = 24 )

# retrieve all emails in the inbox, then sort them from most recently received to oldest (False will give you the reverse). Not strictly necessary, but good to know if order matters for your search
messages = root_folder.Items
messages.Sort("[ReceivedTime]", True)

# Formats the delta_date_time variable
delta_date_time = messages.Restrict("[ReceivedTime] >= '" +delta_date_time.strftime('%m/%d/%Y %I:%M %p')+"'")
# Current date and time 
current_date_time = "Current date and time: " + date_time.strftime('%m/%d/%Y %I:%M %p')

# list emails from Sender over the last 24 hours
emails_list = "" 
count = 0
for message in delta_date_time:
   if message.SenderEmailAddress == "jblakely@barnetproducts.com":
      print(message)


# opens toggle and n-able in default browser 
webbrowser.open("https://track.toggl.com/timer")
webbrowser.open("https://www.n-able.com/product-login")
