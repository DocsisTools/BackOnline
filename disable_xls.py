
from interface_stage import cisco_xls_export_button,casa_xls_export_button

def disable_xls_buttons():
    cisco_xls_export_button.configure(state='disabled')
    casa_xls_export_button.configure(state='disabled')
    cisco_xls_export_button.configure(fg_color='transparent')
    casa_xls_export_button.configure(fg_color='transparent')
    
def enable_cisco_xls_button():
    cisco_xls_export_button.configure(state='normal',fg_color='#4FB140',hover_color='#66C552')
    
def disable_cisco_xls_button():
    cisco_xls_export_button.configure(state='disabled')
    cisco_xls_export_button.configure(fg_color='transparent')
    
def enable_casa_xls_button():
    casa_xls_export_button.configure(state='normal',fg_color='#4FB140',hover_color='#66C552')
    
def disable_casa_xls_button():
    casa_xls_export_button.configure(state='disabled')
    casa_xls_export_button.configure(fg_color='transparent')