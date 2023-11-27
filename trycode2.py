from pywinauto import Application
import time
import subprocess

import win32com.client

from pywinauto.application import Application
import pyautogui


# Start SAP GUI application
# The path with the actual path to your SAP GUI executable
sap_gui_path = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"

# Close any existing SAP GUI windows
app = Application(backend='uia')
try:
    app.connect(title_re=".*SAP Logon.*", timeout=5)
    window = app.window(title_re=".*SAP Logon.*")
    window.close()
    print("Closed existing SAP GUI window.")
except:
    pass  # No existing SAP GUI window found

# Launch SAP GUI
subprocess.Popen([sap_gui_path])
print("SAP GUI has been started.")
# Create an instance of the SAP GUI Scripting COM object
time.sleep(1)
app = Application(backend='uia').connect(title_re=".*SAP Logon.*", timeout=20)
window = app.window(title_re=".*SAP Logon.*")
# Maximize the SAP GUI window
window.maximize()
#print all the elements in the home SAP screen
elements = window.descendants()
list = [element for element in elements if 'P10 (ERP Production)' in element.window_text() or 'Filter Items:' in element.window_text()]
print(f"Found {len(list)} 'P10 (ERP Production)' elements")

if list:
    P10 = list[0]
    print("P10 Element:", P10)
print(P10)

time.sleep(1)

# click on filter items box and put the value P10 and hit enter

filter_item_box=list[2]
filter_item_box.click_input()
time.sleep(1)
#write the  the value P10 inside the box to filter
filter_item_box.type_keys('P10')
time.sleep(1)

# Re-find the element to ensure it reflects the original state
#the new window after adding the value
window = app.window(title_re=".*SAP Logon.*")
#extracting the new elements in the new window
items = window.descendants()
Relist=[]
for item in items:
 if 'P10 (ERP Production)' in item.window_text():
        print(item)
        Relist.append(item)
        print(Relist[0])
        time.sleep(2)
        # double clicking on the server that we want to work on which P10
        Relist[0].double_click_input()
        time.sleep(2)
        break
print(len(Relist))

# set up the connection to automate the work

SapGuiAuto = win32com.client.GetObject("SAPGUI")
print(SapGuiAuto,"success")

    
SapGuiAuto = win32com.client.GetObject("SAPGUI")
if not type(SapGuiAuto) == win32com.client.CDispatch:
    print(SapGuiAuto)
    time.sleep(2)  

# Connect to SAP GUI Scripting Engine
application = SapGuiAuto.GetScriptingEngine
if not type(application) == win32com.client.CDispatch:
    SapGuiAuto = None
print(application)
                        
connection = application.Children(0)
if not type(connection) == win32com.client.CDispatch:
    application = None
    SapGuiAuto = None 
print(connection) 
session = connection.Children(0)
print(session)
if not type(session) == win32com.client.CDispatch:
    connection = None
    application = None
    SapGuiAuto = None 
server_name = session.Info.SystemName
print("Server Name:", server_name)   



# Maximize the window and perform actions on elements
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "amierraf"
time.sleep(3)
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Imation12!"

session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
time.sleep(3)
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10
time.sleep(3)
session.findById("wnd[0]").sendVKey(0)
time.sleep(3)
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("F00005")
time.sleep(1)
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00061"
time.sleep(1)
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "F00005"
time.sleep(1)
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00061")
time.sleep(1)
session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = "zps000000005"
time.sleep(1)
session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").caretPosition = 12
time.sleep(1)
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(1)
session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = "DM-180329-01"
time.sleep(1)
session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").setFocus()
time.sleep(1)
session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").caretPosition = 12
time.sleep(1)
session.findById("wnd[0]/tbar[1]/btn[8]").press()
time.sleep(1)

print(input("press enter "))


