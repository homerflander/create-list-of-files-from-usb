#first install wmi and pypiwin32 packages

import wmi
import os
import io
import datetime
import time

cwd = os.getcwd()

#function for recursive search
def search(scriptparam,depth):
    dirList = os.listdir("./")
    for filename in dirList:
        x=0
        holder=""
        for x in range (0,depth):
            holder=holder+"\t"
        if(depth>0):
            holder = holder + "->" + filename
        else:
            holder=holder+filename
        #print(depth) #was used for troubleshooting depth of subfiles/folders
        if (filename != scriptparam):
            #print(holder)
            FILEOUT.write(holder + "\n")
            if os.path.isdir(filename):
                #print(filename + " is a folder") #was used from troubleshooting if file was a folder or file
                try:
                    os.chdir(filename)
                    depth = depth + 1
                    search(scriptparam, depth)
                    depth = depth - 1
                    os.chdir("..")
                except PermissionError:
                    print("Permission Denied at "+filename)
                    continue

while 1:
    print("STARTING")
    os.chdir(cwd)
    c = wmi.WMI()
    break_var = False
    old_list = ''
    location = ''
    outfile = ''

    #create old list for attached devices
    for drive in c.Win32_DiskDrive():
        old_list = old_list + drive.SerialNumber
        #print(drive)

    #loop until new device is connected
    while break_var == False:
        for drive in c.Win32_DiskDrive():
            print("RUNNING USB CHECK")
            time.sleep(1)
            #if found, get device letter using device serial number with query
            if old_list.find(drive.SerialNumber) == -1:
                    #print(drive)
                    break_var = True
                    #discovered/edited this from https://stackoverflow.com/questions/40134760/identify-drive-letter-of-usb-device-using-python (jgrant)
                    #read up more at https://msdn.microsoft.com/en-us/library/aa394606(v=vs.85).aspx
                    for disk in c.query('SELECT * FROM Win32_DiskDrive WHERE SerialNumber LIKE "'+drive.SerialNumber+'"'):
                        deviceID = disk.DeviceID
                        for partition in c.query('ASSOCIATORS OF {Win32_DiskDrive.DeviceID="' + deviceID + '"} WHERE AssocClass = Win32_DiskDriveToDiskPartition'):
                            for logical_disk in c.query('ASSOCIATORS OF {Win32_DiskPartition.DeviceID="' + partition.DeviceID + '"} WHERE AssocClass = Win32_LogicalDiskToPartition'):
                                print('Drive letter: {}'.format(logical_disk.DeviceID))
                                location=format(logical_disk.DeviceID)

    #create new output file based on time
    outfile = str(datetime.datetime.now())
    outfile=outfile.replace(':', '-', 2)
    outfile=outfile.replace('.', ' ', 1)
    outfile=outfile+".txt"
    FILEOUT=io.open(outfile, "w", encoding='utf-8')
    os.chdir(location)

    #open output.txt file to write the list, excluding the python script name
    search(location,0)
    FILEOUT.close()
    print("ENDING")
