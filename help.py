import customtkinter
from CTkMessagebox import CTkMessagebox

def help_pop():
    msg_main =CTkMessagebox(title = 'HELP',width = 1200,icon='', height=300,option_3='CISCO_HELP',option_2='CASA_HELP',option_1='QUIT', message = '''

Quick Start Guide:

1. Insert Data for Analysis:

1.a. Option 1: Paste output of 'show cable modem' (SCM) into leftmost textbox. Then select the vendor (CISCO/CASA) and corresponding data command (BEFORE/AFTER).

1.b. Option 2: Click on 'BROWSE_CISCO' or 'BROWSE_CASA' (BEFORE/AFTER) to locate the correct .txt file on your local PC. The file will be automatically loaded with the correct values after selection.         
                  ''')

    if msg_main.get() =="CISCO_HELP":
        cisco_pop()
        
    elif msg_main.get() == "CASA_HELP":
        casa_pop()
    
    else:
        msg_main.destroy() 
    
def cisco_pop():
    msg_main =CTkMessagebox(title = 'HELP',width = 1200, height=300,icon='',justify=True,option_3='MAIN_HELP',option_2='CASA_HELP',option_1='QUIT', message = '''
                                                     
2. Choose Analysis Type:

IMPORTANT NOTE: Both 'Before' and 'After' data must be loaded for accurate checks. The analysis won't work if only one side is loaded.

2.a Cisco Functions - Output in the Rightmost Textbox (Except Cisco_Detail):

>>>> CISCO_COMPARE_ALL: 
Compares the state of all devices in before and after files.
Output format: MAC, STATUS_BEFORE, STATUS_AFTER.

>>>> CISCO_ONLINE_BEFORE: 
Displays devices that were in 'Status_Before: Online (pt)' and 'W-Online (pt)', but in different states after.

>>>> CISCO_BEFORE!=AFTER: 
Provides a broader scope, showing all devices that are in two different states, regardless of which states they are.

>>>> CISCO_DETAIL: 
Main analysis function with output in the Leftmost Textbox. It presents detailed analysis in the following format: MAC, STATUS_BEFORE, STATUS_AFTER, INTERFACE_BEFORE, INTERFACE_AFTER. Note: This output is exportable to Excel.

3.  Exporting Data to Excel:

>>>> CISCO_TO_XLS: 
This function exclusively operates on the output of the Cisco_Detail function. It exports the output to "cisco_analysis.xlsx" in the folder where the application is located and opens it automatically.

''')
    
    if msg_main.get() =="MAIN_HELP":
        help_pop()
        
    elif msg_main.get() == "CASA_HELP":
        casa_pop()
    
    else:
        msg_main.destroy() 
    
    
def casa_pop():
    msg_main =CTkMessagebox(title = 'HELP',width = 1200,icon='', height=300,icon_size=(105,105),justify=True,option_3='MAIN_HELP',option_2='CISCO_HELP',option_1='QUIT', message = '''
                            
                            
2.b CASA Functions - Output in the Rightmost Textbox (Except Casa_Detail):

>>>> CASA_COMPARE_ALL:           
Compares the state of all devices in before and after files. Output format: MAC, STATUS_BEFORE, STATUS_AFTER.

>>>> CASA_ONLINE_BEFORE:         
Displays devices that were in 'Status_Before: Online (pt)' and in different status_after.

>>>> CASA_BEFORE!=AFTER:         
Provides a broader scope, showing all devices that are in two different states, regardless of which states they are.

>>>> CASA_DETAIL:                
Main analysis function with output in the Leftmost Textbox. It presents detailed analysis in the following format: MAC, STATUS_BEFORE, STATUS_AFTER, US_INTERFACE_BEFORE, US_INTERFACE_AFTER, DS_INTERFACE_BEFORE, DS_INTERFACE_AFTER. Note: This output is exportable to Excel.

>>>> CASA_BONDING_DIFF & PARTIAL: 
Casa-specific function showing devices that were in US bonding before and in different bonding states after, bonding only one channel, or working in partial-service mode (indicated by '#' at the end of bonding).

3. Exporting Data to Excel:

>>>> CASA_TO_XLS: 
This function exclusively works on the output of Casa_Detail function. It exports the output to "casa_analysis.xlsx" in the folder where the app is located and opens it automatically.

''')
    
    if msg_main.get() =="MAIN_HELP":
        help_pop()
        
    elif msg_main.get() == "CISCO_HELP":
        cisco_pop()
    
    else:
        msg_main.destroy()