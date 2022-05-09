# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client
import datetime
"""
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 6 -> to check inbox folder
inbox = outlook.GetDefaultFolder(6)

#gets the collection of Objects
messages = inbox.Items


message = messages.GetLast()

d = (datetime.date.today() - datetime.timedelta (days=1)).strftime("%d-%m-%y")

while message:
    print(message.SentOn)
    print(message.Subject)
    if message.SentOn == d:
        sjl = message.Subject
        print(message.SentOn)
        print(sjl)
        break
    #message = message.GetPrevious()
"""
def getoutlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook

def retrieve_messages(outlook):
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages = Items.Sort("[ReceivedTime]", False)
    k = 0
    for message in messages:
        print(message.SentOn)
        print(message.Subject)
        k+=1
        if k == 5:
            break

def main():
    outlook = getoutlook()
    retrieve_messages(outlook)

if __name__ == "__main__":
    main()