# Import some COM stuff
from win32com.propsys import propsys
from win32com.shell import shellcon
import pythoncom
import os

def get_org_name(filename, fileProperty):

    #Define the property key we want to read and get a pointer to var holding prop key we want
    propKey = propsys.PSGetPropertyKeyFromName(fileProperty)

    try:

        # Here we are accessing the property store and gettng an interface pointer
        propStore = propsys.SHGetPropertyStoreFromParsingName(filename, None, shellcon.GPS_DEFAULT, propsys.IID_IPropertyStore)
    
        # use the Property key value to "index" into the interface for the interface pointer
        propVal = propStore.GetValue(propKey).GetValue()
        
        
    except:
        propVal = "ERROR"

    
    return propVal


#### --------------- MAIN ------------------ ####
targetDir = "C:\\Users\\jlinebau\\Desktop\\TEST"
dirPath = os.fsencode(targetDir)

# Will process a directory of files - not currently recursive
for file in os.listdir(dirPath):
    fname = os.fsdecode(file)
    fname = targetDir + "\\" + fname

    # Get current directory filename and Original name from file poperties
    propPath = get_org_name(fname, "System.ItemPathDisplay")
    propReal = get_org_name(fname, "System.FileName")
    propOrig = get_org_name(fname, "System.OriginalFileName")
    propParse = get_org_name(fname, "System.ParsingName")
    

    # Lets see if we have a delta in the names that may be suspicious
    if ( (propReal != propOrig) and (propParse != propOrig) and (propOrig != None)):
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n")
        print("#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#")
        print("[<*!*>] ---> **SUSPECT** Filname: %s" %(propReal))
        print("[<*!*>] ---> **REALNAME** OriginalFileName: %s" %(propOrig))
        print("#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#")
        res = get_org_name(fname, "System.ItemType")
        print("[**] ---> ItemType: %s" %(res))
        res = get_org_name(fname, "System.FileDescription")
        print("[**] ---> FileDescription: %s" %(res))
        res = get_org_name(fname, "System.InternalName")
        print("[**] ---> InternalName: %s" %(res))
        res = get_org_name(fname, "System.FileOwner")
        print("[**] ---> FileOwner: %s" %(res))
        res = get_org_name(fname, "System.ComputerName")
        print("[**] ---> ComputerName: %s" %(res))
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n\n")

    elif ((propPath == "ERROR") or (propReal == "ERROR") or (propOrig == "ERROR") or (propParse == "ERROR") or (propOrig == None)):
        print("---------------------------------------------------\n\n")
        print("[<!!>] ---> ERROR - File may not have properties we need: ItemPathDisplay|FileName|OriginalFileName : %s" %(propReal))
        print("[**] ---> FileName: %s" %(propReal))
        res = get_org_name(fname, "System.ItemNameDisplay")
        print("[**] ---> ItemNameDisplay: %s" %(res))
        res = get_org_name(fname, "System.ItemType")
        print("[**] ---> ItemType: %s" %(res))
        res = get_org_name(fname, "System.FileDescription")
        print("[**] ---> FileDescription: %s" %(res))
        res = get_org_name(fname, "System.InternalName")
        print("[**] ---> InternalName: %s" %(res))
        res = get_org_name(fname, "System.FileOwner")
        print("[**] ---> FileOwner: %s" %(res))
        res = get_org_name(fname, "System.ComputerName")
        print("[**] ---> ComputerName: %s" %(res))
        print("---------------------------------------------------\n\n")


    else:
        #otherwise all clear
        print("---------------------------------------------------\n\n")
        print("[**] ---> FNAME MATCH %s: %s\n" %(propOrig, propPath))
        print("---------------------------------------------------\n\n")


    
        
    

    











# Create data to possibly change the value
#new value = propsys.PROPVARIANTType(["test", "test"], pythoncom.VT_VECTOR | pythoncom.VT_BSTR)
##SetValue(key, something else)
#Commit()