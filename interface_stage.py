import customtkinter
import tkinter

customtkinter.set_appearance_mode('Dark')
customtkinter.set_default_color_theme('dark-blue')

root = customtkinter.CTk()
root.title('BackOnline v1.0')
root.geometry('1450x570')
root.iconbitmap('testico.ico')


root.configure(background='black')


def create_ui(root):

    scm_before_input=customtkinter.CTkTextbox(root,width=750,height=300,font=('Consolas',12),corner_radius=5,border_spacing=5,wrap='none')
    scm_before_input.grid(column=0,row=0,padx=2.5,columnspan=1)

    scm_before_data=customtkinter.CTkTextbox(root,width=150,height=300,font=('Consolas',12),corner_radius=5,border_spacing=5,wrap='none')
    scm_before_data.grid(column=1,row=0,padx=2.5,columnspan=1)

    scm_after_data=customtkinter.CTkTextbox(root,width=150,height=300,font=('Consolas',12),corner_radius=5,border_spacing=5,wrap='none')
    scm_after_data.grid(column=2,row=0,padx=2.5,columnspan=1)

    query_output_box=customtkinter.CTkTextbox(root,width=370,height=300,font=('Consolas',12),corner_radius=5,border_spacing=5,wrap='none')
    query_output_box.grid(column=3,row=0,padx=2.5,columnspan=1)

    return scm_before_input, scm_before_data, scm_after_data, query_output_box

def create_button_frame(root):
    
    button_frame=customtkinter.CTkFrame(root)
    button_frame.grid(column=0,row=1,columnspan=4,rowspan=4,sticky='N',pady=20)
    return button_frame

create_ui(root)
button_frame = create_button_frame(root)

def btn_placeholder():
    print('bobnun')

def create_buttons_row_1(button_frame):
    
    cisco_scm_before_button=customtkinter.CTkButton(button_frame,text="CISCO_SCM_BEFORE",
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#00A0DA',hover_color='#33B8FF')

    cisco_scm_before_button.grid(row=0,column=0)

    cisco_scm_after_button=customtkinter.CTkButton(button_frame,text="CISCO_SCM_AFTER" ,                                                   
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#00A0DA',hover_color='#33B8FF')
    
    cisco_scm_after_button.grid(row=0,column=1,padx=15,pady=2)

    cisco_browse_scm_before_button=customtkinter.CTkButton(button_frame,text="BROWSE_CISCO_BEFORE",                                                    
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#00A0DA',hover_color='#33B8FF')
    
    cisco_browse_scm_before_button.grid(row=0,column=2,padx=15,pady=2)

    cisco_browse_scm_after_button=customtkinter.CTkButton(button_frame,text="BROWSE_CISCO_AFTER" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#00A0DA',hover_color='#33B8FF')
    
    cisco_browse_scm_after_button.grid(row=0,column=3,padx=15,pady=2)

    
    help_button=customtkinter.CTkButton(button_frame,text="HELP_BUTTON" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#394457',hover_color='#4D576E',text_color='yellow')
    help_button.grid(row=0,column=5,padx=5,pady=5)

    return cisco_scm_before_button,cisco_scm_after_button,cisco_browse_scm_before_button,cisco_browse_scm_after_button,help_button



def create_buttons_row_2(button_frame):

    scm_cisco_compare_all_button=customtkinter.CTkButton(button_frame,text="CISCO_COMPARE_ALL" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#054FA8',hover_color='#1A6ABD')
    
    scm_cisco_compare_all_button.grid(row=1,column=0,padx=5,pady=5)

    scm_cisco_onl_before_button=customtkinter.CTkButton(button_frame,text="CISCO_ONLINE_BEFORE" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#054FA8',hover_color='#1A6ABD')
    
    scm_cisco_onl_before_button.grid(row=1,column=1,padx=5,pady=5)

    scm_cisco_state_different_button=customtkinter.CTkButton(button_frame,text="CISCO_BEFORE!=AFTER" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#054FA8',hover_color='#1A6ABD')
    
    scm_cisco_state_different_button.grid(row=1,column=2,padx=5,pady=5)
    
    scm_cisco_detail_button=customtkinter.CTkButton(button_frame,text="CISCO_DETAIL" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#054FA8',hover_color='#1A6ABD')
    #F6821D fg_color='#F6821D'
    
    scm_cisco_detail_button.grid(row=1,column=3,padx=5,pady=5)
    
    cisco_xls_export_button=customtkinter.CTkButton(button_frame,text="CISCO_TO_XLSX" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#4FB140',hover_color='#66C552')
    
    cisco_xls_export_button.grid(row=1,column=5,padx=15,pady=5)

    return scm_cisco_compare_all_button,scm_cisco_onl_before_button,scm_cisco_state_different_button,scm_cisco_detail_button,cisco_xls_export_button




def create_buttons_row_3(button_frame):
    
    scm_casa_scm_before_button=customtkinter.CTkButton(button_frame,text="CASA_SCM_BEFORE" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#EF820D',hover_color='#FF9D40')
    
    scm_casa_scm_before_button.grid(row=2,column=0,padx=5,pady=5)

    scm_casa_scm_after_button=customtkinter.CTkButton(button_frame,text="CASA_SCM_AFTER" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#EF820D',hover_color='#FF9D40')
    
    scm_casa_scm_after_button.grid(row=2,column=1,padx=5,pady=5)

    scm_casa_scm_browser_before_button=customtkinter.CTkButton(button_frame,text="BROWSE_CASA_BEFORE" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#EF820D',hover_color='#FF9D40')
    
    scm_casa_scm_browser_before_button.grid(row=2,column=2,padx=5,pady=5)

    scm_casa_scm_browser_after_button=customtkinter.CTkButton(button_frame,text="BROWSE_CASA_AFTER" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#EF820D',hover_color='#FF9D40')
    
    scm_casa_scm_browser_after_button.grid(row=2,column=3,padx=5,pady=5)


    return scm_casa_scm_before_button, scm_casa_scm_after_button, scm_casa_scm_browser_before_button, scm_casa_scm_browser_after_button,




def create_buttons_row_4(button_frame):
    casa_scm_compare_all_button=customtkinter.CTkButton(button_frame,text="CASA_COMPARE_ALL" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#CC461B',hover_color='#E55D2F')
    
    casa_scm_compare_all_button.grid(row=3,column=0,padx=5,pady=5)

    casa_scm_online_before_button=customtkinter.CTkButton(button_frame,text="CASA_ONLINE_BEFORE" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#CC461B',hover_color='#E55D2F')
    
    casa_scm_online_before_button.grid(row=3,column=1,padx=5,pady=5)

    casa_scm_different_state_button=customtkinter.CTkButton(button_frame,text="CASA_BEFORE!=AFTER" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#CC461B',hover_color='#E55D2F')
    
    casa_scm_different_state_button.grid(row=3,column=2,padx=5,pady=5)

    casa_detail_button=customtkinter.CTkButton(button_frame,text="CASA_DETAIL" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#CC461B',hover_color='#E55D2F')
    
    casa_detail_button.grid(row=3,column=3,padx=5,pady=5)

    casa_scm_bonding_compare_button=customtkinter.CTkButton(button_frame,text="CASA_BONDING_DIFF & PARTIAL" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#CC461B',hover_color='#E55D2F')
    
    casa_scm_bonding_compare_button.grid(row=3,column=4,padx=5,pady=5)

    
    casa_xls_export_button=customtkinter.CTkButton(button_frame,text="CASA_TO_XLSX" ,
                                                    command=btn_placeholder,
                                                    width=200,height=45,
                                                    corner_radius=15,
                                                    font = ('Consolas',14),
                                                    fg_color='#4FB140',hover_color='#66C552')
    
    casa_xls_export_button.grid(row=3,column=5,padx=5,pady=5)
    
    return casa_scm_compare_all_button,casa_scm_online_before_button,casa_scm_different_state_button,casa_scm_bonding_compare_button,casa_detail_button,casa_xls_export_button




create_buttons_row_1(button_frame)
create_buttons_row_2(button_frame)
create_buttons_row_3(button_frame)
create_buttons_row_4(button_frame)


scm_before_input, scm_before_data, scm_after_data, query_output_box = create_ui(root)

button_frame = create_button_frame(root)

cisco_scm_before_button,cisco_scm_after_button,cisco_browse_scm_before_button,cisco_browse_scm_after_button,help_button = create_buttons_row_1(button_frame)
scm_cisco_compare_all_button,scm_cisco_onl_before_button,scm_cisco_state_different_button,scm_cisco_detail_button,cisco_xls_export_button = create_buttons_row_2(button_frame)
scm_casa_scm_before_button, scm_casa_scm_after_button, scm_casa_scm_browser_before_button, scm_casa_scm_browser_after_button = create_buttons_row_3(button_frame)
casa_scm_compare_all_button,casa_scm_online_before_button,casa_scm_different_state_button,casa_scm_bonding_compare_button,casa_detail_button,casa_xls_export_button = create_buttons_row_4(button_frame)

cisco_status_label = customtkinter.CTkLabel(master=button_frame,text='', font=('Consolas', 14),corner_radius=15)
cisco_status_label.grid(column=4,row=0)

casa_status_label = customtkinter.CTkLabel(master=button_frame,text='', font=('Consolas', 14),corner_radius=15)
casa_status_label.grid(column=4,row=2)

from cisco_functions import *
from casa_functions import *
from help import *
cisco_scm_before_button.configure(command=input_scm_before)
cisco_scm_after_button.configure(command=input_scm_after)
cisco_browse_scm_before_button.configure(command=cisco_scm_browse_before)
cisco_browse_scm_after_button.configure(command=cisco_scm_browse_after)
help_button.configure(command=help_pop)
cisco_xls_export_button.configure(command=export_xl_cisco)

scm_cisco_compare_all_button.configure(command=cisco_join_compare)
scm_cisco_onl_before_button.configure(command=cisco_onl_diff_compare)
scm_cisco_state_different_button.configure(command=cisco_state_diff_compare)
scm_cisco_detail_button.configure(command=cisco_scm_detail)

scm_casa_scm_before_button.configure(command=casa_scm_before)
scm_casa_scm_after_button.configure(command=casa_scm_after)
scm_casa_scm_browser_before_button.configure(command=casa_scm_browse_pre)
scm_casa_scm_browser_after_button.configure(command=casa_scm_browse_posle)
casa_xls_export_button.configure(command=export_xl_casa)

casa_scm_compare_all_button.configure(command=casa_compare_all)
casa_scm_online_before_button.configure(command=casa_compare_onl_pre)
casa_scm_different_state_button.configure(command=casa_compare_different_state)
casa_scm_bonding_compare_button.configure(command=casa_compare_bonding)
casa_detail_button.configure(command=casa_compare_all_detail)

cisco_xls_export_button.configure(state='disabled')
casa_xls_export_button.configure(state='disabled')
cisco_xls_export_button.configure(fg_color='transparent')
casa_xls_export_button.configure(fg_color='transparent')



black = 1

def choose_mode(): 
    global black
    if black == 1:
        customtkinter.set_appearance_mode('Light')
        black = 0
    else:
        customtkinter.set_appearance_mode('Dark')
        black = 1


def light_dark():
    toggle_mode=customtkinter.CTkButton(button_frame,text="LIGHT/DARK_MODE" ,
                                                        command=choose_mode,
                                                        width=200,height=45,
                                                        corner_radius=15,
                                                        font = ('Consolas',14),
                                                        fg_color='#394457',hover_color='#4D576E',text_color='yellow')
    toggle_mode.grid(row=2,column=5,padx=5,pady=5)



light_dark()
