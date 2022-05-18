from cgitb import text
import win32com.client
import os
import time
import datetime as dt

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
root_folder = mapi.Folders['Support'].Folders['Inbox']

# List the current time in 12 hr string format and replacing ":" and "/" with "_"
current_time = str(dt.datetime.now().strftime("%m/%d/%Y-%I:%M %p"))
current_time = current_time.replace(':', '_').replace('/', '_')

# this is set to the current time
date_time = dt.datetime.now()
#This is set to 8 hours ago; you can change timedelta's argument to whatever you want it to be
delta_date_time = dt.datetime.now() - dt.timedelta(hours = 24 )

# retrieve all emails in the inbox, then sort them from most recently received to oldest (False will give you the reverse). Not strictly necessary, but good to know if order matters for your search
messages = root_folder.Items
messages.Sort("[ReceivedTime]", True)

# Formats the delta_date_time variable
delta_date_time = messages.Restrict("[ReceivedTime] >= '" +delta_date_time.strftime('%m/%d/%Y %I:%M %p')+"'")
# Current date and time 
current_date_time = "Current date and time: " + date_time.strftime('%m/%d/%Y %I:%M %p')

# Iterating through the emails restricted by time. Joining each email by a line break
# Count is the total number of emails containing the specified requirments set
emails_list = "" 
count = 0
for message in delta_date_time:
   if message.subject.startswith("Error"):
      count += 1
      emails_list += ''.join(message.subject + ": " + message.SenderEmailAddress + ": " + message.ReceivedTime.strftime('%I:%M %p') + '\n')
emails_list = emails_list + "\nTotal number: " + str(count)

# Writes the contents of email_list to a folder on the desktop called Scrubber. The file is titled the current date and time
text_file = open("C:\\Users\\MatthewCrow\\Desktop\\Scrubber\\{0}.txt".format(current_time), "w")
text_file.write(emails_list)
text_file.close()

print('\n' + current_date_time + '\n')
print(emails_list)
