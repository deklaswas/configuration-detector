import os
import platform
import tkinter as tk
from tkinter import TclError, ttk
from tkinter import font
from tkinter import filedialog

import datetime

from win32api import GetFileVersionInfo



# fpath - insert the file path as parameter
# returns array of ['ProductName','CompanyName','ProductVersion']
def getFileProperties(fpath):
    #all the property names that we are interested in
    detailNames = ('ProductName', 'CompanyName', 'ProductVersion')
    strInfo = {}

    try:
        # \VarFileInfo\Translation returns list of available (language, codepage) pairs that can be used to retreive string info. We are using only the first pair.
        lang, codepage = GetFileVersionInfo(fpath, '\\VarFileInfo\\Translation')[0]

        # for each of the properties that we are interested in
        for detail in detailNames:
            # get the path to the property in the file
            strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, detail)

            # read it and put it into our return array
            strInfo[detail] = GetFileVersionInfo(fpath, strInfoPath)
    except:
        pass

    return strInfo



def create_input_frame(container):

    frame = ttk.Frame(container)

    # grid layout for the input frame
    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(0, weight=3)

    myFont = font.Font(family='Courier', size=25, weight='bold')

    # Find what
    tk.Label(frame, text='Software Configuration\nDetector for Windows', font=myFont).pack(side=tk.TOP)

    for widget in frame.winfo_children():
        widget.grid(padx=5, pady=5)

    return frame


def create_lower_frame(container,f):
    frame = ttk.Frame(container)

    myFont = font.Font(family='Courier', size=12, weight='bold')

    frame.columnconfigure(0, weight=4)
    frame.columnconfigure(1, weight=1)


    side_frame = create_side_frame(frame,f)
    side_frame.grid(column=0, row=0, sticky=tk.W) #poo

    def generate_info():
        # get array of info
        string_path = f.get()


        # make sure file path exists
        if not os.path.exists(string_path):
            #print('File path not found!')
            f.set('File path not found!')

        # make sure file is an application
        elif os.path.splitext(string_path)[1] != ".exe":
            f.set('Please input an application!')
        
        #good to go
        else:
            f.set('')

            prop_array = getFileProperties(string_path)

            software_name = prop_array['ProductName']
            windows_version = platform.system()+" "+platform.version()
            software_version = prop_array['ProductVersion']
            manufacturer = prop_array['CompanyName']

            info_window = tk.Toplevel()


            info_window.title("File Properties Report")
            info_window.geometry("450x300")
            info_window.resizable(False, False)

            textBox = tk.Text(info_window, height = 20, width = 52, highlightbackground='black')
            textBox.pack()
            textBox.config(font =("Courier", 12))

            hline = "============================================"

            report_data = "REPORT- \nWINDOWS SOFTWARE CONFIGURATION DETECTOR\n"+hline+"\n"
            report_data += string_path+"\n"+hline+"\n\n"
            report_data += "  Software Name:     "+software_name+"\n"
            report_data += "  Software Version:  "+software_version+"\n"
            report_data += "  Manufacturer:      "+manufacturer+"\n\n"
            report_data += "  Windows Version:   "+windows_version+"\n\n"
            report_data += hline+"\n"+"Report generated on "+datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


            textBox.insert(tk.END, report_data)

    tk.Button(frame, text='Make\nReport', command=generate_info,
            bg='#CC5200', fg='#ffffff', font=myFont
            ).grid(column=1, row=0, sticky=tk.E)

    for widget in frame.winfo_children():
        widget.grid(padx=25, pady=20)

    return frame

def create_side_frame(container,f):
    frame = ttk.Frame(container)
    frame.pack(expand=True, fill=tk.X)

    myFont = font.Font(family='Courier', size=14, weight='bold')
    
    system_name = platform.system()      # Just the system/OS name, e.g., 'Windows'
    system_version = platform.version()     # System's release version
    system_config = system_name+" "+system_version
    tk.Label(frame, text='Windows Version: '+system_config).grid(column=0, row=0, sticky=tk.NW, pady=20)

    path_frame = create_path_frame(frame,f)
    path_frame.grid(column=0, row=1, sticky=tk.E) #poo

    #for widget in frame.winfo_children():
    #    widget.grid(padx=5, pady=5)

    return frame

def create_path_frame(container,f):
    frame = ttk.Frame(container)
    
    myFont = font.Font(family='Courier', size=14, weight='bold')
    
    def select_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            f.set(file_path)
    
    tk.Button(frame, text='Select File', command=select_file,
            bg='#0052CC', fg='#ffffff', font=myFont
            ).grid(column=0, row=0)
    
    tk.Label(frame, textvariable=f).grid(column=1, row=0)

    

    return frame


def create_main_window():
    root = tk.Tk()
    root.title('Replace')
    root.geometry("700x400")
    root.resizable(0, 0)
    root.configure(bg='white')
    try:
        # windows only (remove the minimize/maximize button)
        root.attributes('-toolwindow', True)
    except TclError:
        print('Not supported on your platform')

    # layout on the root window
    root.rowconfigure(0, weight=3)
    root.rowconfigure(1, weight=1)


    # string to display file path when selected or show error
    file_path = tk.StringVar()

    input_frame = create_input_frame(root)
    input_frame.grid(column=0, row=0)

    button_frame = create_lower_frame(root,file_path)
    button_frame.grid(column=0, row=1, columnspan=500)

    root.mainloop()


if __name__ == "__main__":
    create_main_window()
