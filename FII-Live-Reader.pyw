import tkinter as tk
from display import FIIConsole

root = tk.Tk()

console = FIIConsole(root, manual=False)

console.refresh()

root.mainloop()