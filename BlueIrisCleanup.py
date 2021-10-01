import winreg
import os
from winreg import *
from os import listdir
from os.path import isfile, join
from os import walk
import glob
import ctypes
import platform
import sys
import wmi
from datetime import datetime, time, timedelta
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

free_space_minimum = 50

yesterday_midnight = (datetime.combine(datetime.today(), time.min)) - timedelta(days=1)

NewClipsKey = r"SOFTWARE\Perspective Software\Blue Iris\clips\folders\0"
StoredClipsKey = r"SOFTWARE\Perspective Software\Blue Iris\clips\folders\1"
AlertsClipsKey = r"SOFTWARE\Perspective Software\Blue Iris\clips\folders\2"
OptionsKey = r"SOFTWARE\Perspective Software\Blue Iris\options"

newDict = {}
storedDict = {}
alertsDict = {}

actions_taken = []

sysname = ""

def send_email(cctv_server, actions_taken):

    me = "blueiris-script@lindisfarne.nsw.edu.au"
    you = "itdepartment@lindisfarne.nsw.edu.au"
    bcc = ""

    you_list = you.split(",")
    bcc_list = bcc.split(",")

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Blue Iris Cleanup Script - " + cctv_server
    msg['From'] = formataddr((str(Header('Blue Iris Cleanup Script', 'utf-8')), me))
    msg['To'] = you

    email_text = "" #<p><span style=\"font-weight: 400;\">The following actions were taken:<br><br>"
    
    for action in actions_taken: 
        email_text = email_text + action + "<br>"
        
    email_text = email_text + "</span></p>"

    email_body = MIMEText(email_text, 'html')

    msg.attach(email_body)

    s = smtplib.SMTP('smtp.lindisfarne.nsw.edu.au', 25)

    email_result = s.sendmail(me, you_list + bcc_list, msg.as_string())

    s.quit()

def getFreeSpace(drive_letter):
    c = wmi.WMI ()
    retval = 0
    for d in c.Win32_LogicalDisk():
        if (d.Caption.upper() == drive_letter.upper()):
            retval = d.FreeSpace
    return retval

# Get the System Name
key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, OptionsKey, 0, winreg.KEY_READ)
try:
    count = 0
    while 1:
        name, value, type = winreg.EnumValue(key, count)
        if (name == "sysname"):
            sysname = value
        count = count + 1
except WindowsError:
    pass    

# Get the New information
key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, NewClipsKey, 0, winreg.KEY_READ)
try:
    count = 0
    while 1:
        name, value, type = winreg.EnumValue(key, count)
        newDict[name] = value  
        count = count + 1
except WindowsError:
    pass
    
# Get the Stored information
key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, StoredClipsKey, 0, winreg.KEY_READ)
try:
    count = 0
    while 1:
        name, value, type = winreg.EnumValue(key, count)
        storedDict[name] = value  
        count = count + 1
except WindowsError:
    pass

# Get the Alerts information
key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, AlertsClipsKey, 0, winreg.KEY_READ)
try:
    count = 0
    while 1:
        name, value, type = winreg.EnumValue(key, count)
        alertsDict[name] = value  
        count = count + 1
except WindowsError:
    pass    

# Enumerate files inside each folder
newFiles = glob.glob(newDict["path"] + "\*")
storedFiles = glob.glob(storedDict["path"] + "\*")
alertsFiles = glob.glob(alertsDict["path"] + "\*")

# Get the drive letters for each folder
newFolderDriveLetter = newDict["path"].split(":")[0]
storedFolderDriveLetter = storedDict["path"].split(":")[0]
alertsFolderDriveLetter = alertsDict["path"].split(":")[0]

# Get the free space on each folder's drive in Gigabytes
newFolderFreeSpace = int((str(getFreeSpace(newFolderDriveLetter + ":")))) / 1024 / 1024 / 1024
storedFolderFreeSpace = int((str(getFreeSpace(storedFolderDriveLetter + ":")))) / 1024 / 1024 / 1024
alertsFolderFreeSpace = int((str(getFreeSpace(alertsFolderDriveLetter + ":")))) / 1024 / 1024 / 1024

# New folder
# If free space is less than 5GB, do something
if (newFolderFreeSpace < free_space_minimum):
    print("New folder (" + newDict["path"] + ") free space is less than " + str(free_space_minimum) + "GB, seeing what we can do...")
    actions_taken.append("New folder (" + newDict["path"] + ") free space is less than " + str(free_space_minimum) + "GB! (" + str(newFolderFreeSpace) + "GB)<br>The following actions were taken:<br><br>")
    # Move the oldest file to the Stored folder
    print("Number of new files: " + str(len(newFiles)))
    actions_taken.append("Number of new files: " + str(len(newFiles)) + "<br>")
    print("Found files:")
    actions_taken.append("Found files:<br>")
    for file in newFiles:

        st=os.stat(file)    
        mtime=st.st_mtime
 
        print(file + "(" + str(mtime) + ")")
        actions_taken.append(file + "(" + str(mtime) + ")<br>")

        if (float(mtime) < float(yesterday_midnight.timestamp())):
            print("File '" + file + "' is older than today, moving to Stored folder")
            print("Moving '" + file + "' to '" + storedDict["path"])
            try:
                shutil.move(file, storedDict["path"] + "\\")
                print(" - Successfully moved the file")
                actions_taken.append("Moved New file " + file + " to the Stored folder")
            except:
                print("There was an error moving file '" + file + "'")
                actions_taken.append("Error moving New file " + file + " to the Stored folder")
        else:
            print("File '" + file + "' has today's timestamp, ignoring")
            pass
else:
    print("New folder free space is greater than " + str(free_space_minimum) + "GB, don't need to do anything")
    
# Stored folder
# If free space is less than 5GB, do something
if (len(actions_taken) == 0):
    if (newFolderFreeSpace < free_space_minimum):
        print("New folder free space is less than " + str(free_space_minimum) + "GB, seeing what we can do...")
        actions_taken.append("Stored folder free space is less than " + str(free_space_minimum) + "GB! (" + str(storedFolderFreeSpace) + "GB)<br>The following actions were taken:<br><br>")
        # Delete the oldest file
        oldest_file = ""
        oldest_file_age = 0
        
        for file in storedFiles:
            st=os.stat(file)    
            mtime=st.st_mtime
            if (float(mtime) < float(yesterday_midnight.timestamp())):
                if (float(mtime) < oldest_file_age or oldest_file_age == 0):
                    oldest_file = file
                    oldest_file_age = mtime
            else:
                pass

        #print(oldest_file)
        #print(oldest_file_age)

        if (oldest_file != ""):
            print("File '" + oldest_file + "' is older than today, and is the oldest file in the Stored folder")
            print("Deleting '" + oldest_file + "'")
            try:
                os.remove(oldest_file)
                print(" - Successfully deleted the file")
                actions_taken.append("Deleted Stored file " + file)
            except:
                print("There was an error deleting the file '" + file + "'")
                actions_taken.append("Error deleting Stored file " + file)

    else:
        print("Stored folder free space is greater than " + str(free_space_minimum) + "GB, don't need to do anything")

# Alerts folder
# If free space is less than 5GB, do something
if (len(actions_taken) == 0):
    if (alertsFolderFreeSpace < free_space_minimum):
            print("Alerts folder free space is less than " + str(free_space_minimum) + "GB, seeing what we can do...")
            actions_taken.append("Alerts folder free space is less than " + str(free_space_minimum) + "GB! (" + str(alertsFolderFreeSpace) + "GB)<br>The following actions were taken:<br><br>")
            # Delete the oldest file
            oldest_file = ""
            oldest_file_age = 0
            
            for file in alertsFiles:
                st=os.stat(file)    
                mtime=st.st_mtime
                if (float(mtime) < float(yesterday_midnight.timestamp())):
                    if (float(mtime) < oldest_file_age or oldest_file_age == 0):
                        oldest_file = file
                        oldest_file_age = mtime
                else:
                    pass

            #print(oldest_file)
            #print(oldest_file_age)

            if (oldest_file != ""):
                print("File '" + oldest_file + "' is older than today, and is the oldest file in the Alerts folder")
                print("Deleting '" + oldest_file + "'")
                try:
                    os.remove(oldest_file)
                    print(" - Successfully deleted the file")
                    actions_taken.append("Deleted Alert file " + file)
                except:
                    print("There was an error deleting the file '" + file + "'")
                    actions_taken.append("Error deleting Alert file " + file)
    else:
        print("Alerts folder free space is greater than " + str(free_space_minimum) + "GB, don't need to do anything")                    

if (len(actions_taken) > 0):
    send_email(sysname, actions_taken)