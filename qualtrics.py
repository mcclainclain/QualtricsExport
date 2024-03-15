from customtkinter import *
from tkinter import PhotoImage, messagebox
from api import *
from PIL import Image, ImageTk
from tkcalendar import Calendar, DateEntry

import sys
from os import path, environ
import subprocess



api_default = True
path_default = True


root = CTk()




def pull_data(file_path_entry, selected_key, api_entry, start_date, end_date):

    file_path = "N:/DoIT/US/Metrics/Core Monthly Report/Qualtrics/" if path_default else file_path_entry.get()
    api_key = selected_key.get() if api_default else api_entry.get()


    # Call the API
    run(api_key, file_path, root, start_date.get(), end_date.get())

    # Save the key to the keys file
    if api_key not in get_keys():
        save_key(api_key)

    # Close after counting down for 5 seconds
    closeLabel = CTkLabel(root, text="Closing in 5 seconds...")
    closeLabel.grid(row=14, padx=10, pady=5)
    for i in range(5, 0, -1):
        closeLabel.configure(text="Closing in {} seconds...".format(i))
        root.update()
        root.after(1000)

    # Print the file path
    closeLabel.configure(text="Opening the folder...")
    root.update()
    root.after(1000)

    # Open the file path directory in the file explorer
    if sys.platform == "win32":
        os.startfile(file_path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", file_path])

    root.destroy()


def api_sel(label1, entry1, label2, entry2):
    global api_default
    api_default = not api_default
    
    # Hide label and entry if default is selected

    if api_default:
        label2.grid_forget()
        entry2.grid_forget()
        label1.grid(row = 2, padx=10, pady=5)
        entry1.grid(row = 3, padx=10, pady=0)

    else:
        label1.grid_forget()
        entry1.grid_forget()
        label2.grid(row = 2, padx=10, pady=5)
        entry2.grid(row = 3, padx=10, pady=0)
    
def path_sel(entry, label, browseButton):
    global path_default
    path_default = not path_default

    if path_default:
        entry.grid_forget()
        label.grid(row=6, padx=10, pady=0)
        browseButton.grid_forget()
    else:
        label.grid_forget()
        entry.grid(row=6, padx=10, pady=0)
        browseButton.grid(row=7, padx=10, pady=0)

def get_keys():
    if sys.platform == "darwin":
        from AppKit import NSSearchPathForDirectoriesInDomains
        appdata = path.join(NSSearchPathForDirectoriesInDomains(14, 1, True)[0], "qualtrics")
    elif sys.platform == "win32":
        appdata = path.join(environ["LOCALAPPDATA"], "qualtrics")
    else:
        appdata = path.expanduser(path.join("~", ".qualtrics"))

    if not path.exists(appdata):
        os.makedirs(appdata)

    if not path.exists(path.join(appdata, "keys.txt")):
        open(path.join(appdata, "keys.txt"), "w").close()
    
    keys = []
    with open(path.join(appdata, "keys.txt"), "r") as f:
        keys = f.readlines()
        keys = [key.strip() for key in keys]
    return keys

def save_key(key):
    if sys.platform == "darwin":
        from AppKit import NSSearchPathForDirectoriesInDomains
        appdata = path.join(NSSearchPathForDirectoriesInDomains(14, 1, True)[0], "qualtrics")
    elif sys.platform == "win32":
        appdata = path.join(environ["LOCALAPPDATA"], "qualtrics")
    else:
        appdata = path.expanduser(path.join("~", ".qualtrics"))
    
    with open(path.join(appdata, "keys.txt"), "a") as f:
        f.write(key + "\n")
    
def gui():
    
    api_label = CTkLabel(root, text="Enter your qualtrics API token: ")
    api_entry = CTkEntry(master=root, width=250, placeholder_text="Enter your API token here...")

    key_history_label = CTkLabel(root, text="Select your API token from the dropdown: ")
    selected_key = StringVar()
    selected_key.set(get_keys()[0] if len(get_keys()) > 0 else "No keys found. Enter a new key.")
    key_history_dropdown = CTkOptionMenu(master=root, variable=selected_key, values = get_keys())  


    default_checkbox = CTkSwitch(root, text="Used saved API token?", command=lambda: api_sel(key_history_label, key_history_dropdown, api_label, api_entry))
    default_checkbox.select()
    default_checkbox.grid(row=1, padx=10, pady=10)

    key_history_label.grid(row=2, padx=10, pady=5)
    key_history_dropdown.grid(row=3, padx=10, pady=0)



    CTkLabel(root, text = "").grid(row=4)

    path_check = CTkSwitch(root, text="Use default path?", command=lambda: path_sel(file_path_entry, default_path_label, browseButton))
    path_check.select()
    path_check.grid(row=5, padx=10, pady=0)
    

    default_path_label = CTkLabel(
        root, text="Default: N:/DoIT/US/Metrics/Core Monthly Report/Qualtrics/")
    default_path_label.grid(row=6, padx=10, pady=0)


    file_path_entry = CTkEntry(root, width=250, placeholder_text="Enter a path here...")

    # Create a browse button
    browseButton = CTkButton(root, text="Browse", command=lambda: file_path_entry.insert(0, filedialog.askdirectory()))

    
    # Create a label widget
    label = CTkLabel(root, text="Enter the start and end dates:")
    label.grid(row=9, padx=10, pady=5)

    start_date_var = StringVar()
    end_date_var = StringVar()

    # Create a date picker
    start_date = DateEntry(root, width=25, background='#36719f',
                     foreground='white', borderwidth=2, font = ("Arial", 25), date_pattern="mm/dd/yyyy", textvariable=start_date_var)
    start_date.grid(row=10, padx=10, pady=10)

    end_date = DateEntry(root, width=25, background='#36719f',
                     foreground='white', borderwidth=2, font = ("Arial", 25), date_pattern="mm/dd/yyyy" , textvariable=end_date_var)
    end_date.grid(row=11, padx=10, pady=10)

    # Create a button widget
    enterButton = CTkButton(root, text="Generate", command=lambda: pull_data(file_path_entry, selected_key, api_entry, start_date_var, end_date_var))
    enterButton.grid(row=12, padx=10, pady=10)





# When user tries to close the window, ask if they are sure


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()





def main():
    

    root.title("Qualtrics Retrieval")
    root.geometry("600x600")
    root.columnconfigure(0, weight=1)

    gui()

    root.protocol("WM_DELETE_WINDOW", on_closing)


    root.mainloop()

if __name__ == "__main__":
    main()