import tkinter as tk
from tkinter import ttk
import scrape
import datetime
import threading

def extract_data():

    data_map = scrape.get_data()
    
    sorted_data = sorted(data_map.items(), key=lambda x: x[1], reverse=True)
    return sorted_data

root = tk.Tk()
root.title("Script Frequency")

lst_updt_var = tk.StringVar(value="Last Updated: ")
last_updated = tk.Label(root, textvariable=lst_updt_var)
last_updated.pack()

columns = ('#1', '#2')
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.heading('#1', text='Company')
tree.heading('#2', text='Count')
tree.pack()

def map_data():
    extracted_data = extract_data()
    tree.delete(*tree.get_children())
    for line in extracted_data:
        tree.insert("", tk.END, value=line)
    now = datetime.datetime.now().strftime("%H:%M:%S")

    lst_updt_var.set("Last Updated: " + now)
    

def refresh():
    threading.Thread(target=map_data, daemon=True).start()
    root.after(120000, refresh)

refresh()

root.mainloop()