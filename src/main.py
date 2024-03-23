import tkinter as tk
import tkinter.ttk as ttk
import os.path
import subprocess

_location = os.path.dirname(__file__)

class Toplevel1:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        top.geometry("600x450+468+170")
        top.minsize(120, 1)
        top.maxsize(1540, 941)
        top.resizable(1,  1)
        top.title("NPO Report Assistant")
        top.configure(background="#919191")
        top.configure(highlightbackground="#919191")
        top.configure(highlightcolor="white")

        self.top = top

        self.menubar = tk.Menu(top, font="TkMenuFont", bg='#919191', fg='white')
        top.configure(menu=self.menubar)

        self.Button1 = tk.Button(self.top, command=self.connect_to_database)
        self.Button1.place(relx=0.05, rely=0.133, height=26, width=67)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="black")
        self.Button1.configure(background="#919191")
        self.Button1.configure(disabledforeground="#adadad")
        self.Button1.configure(foreground="white")
        self.Button1.configure(highlightbackground="#919191")
        self.Button1.configure(highlightcolor="white")
        self.Button1.configure(text='''Update''')

        self.TSeparator1 = ttk.Separator(self.top)
        self.TSeparator1.place(relx=0.2, rely=0.0,  relheight=0.311)
        self.TSeparator1.configure(orient="vertical")

        self.TSeparator2 = ttk.Separator(self.top)
        self.TSeparator2.place(relx=0.0, rely=0.311,  relwidth=0.2)

        self.Label1 = tk.Label(self.top)
        self.Label1.place(relx=0.05, rely=0.044, height=31, width=54)
        self.Label1.configure(anchor='w')
        self.Label1.configure(background="#919191")
        self.Label1.configure(foreground="white")
        self.Label1.configure(text='''Database''')

        self.Label2 = tk.Label(self.top)
        self.Label2.place(relx=0.1, rely=0.422, height=31, width=94)
        self.Label2.configure(anchor='w')
        self.Label2.configure(background="#919191")
        self.Label2.configure(foreground="white")
        self.Label2.configure(text='''4G Parameters''')

        self.Button2 = tk.Button(self.top)
        self.Button2.place(relx=0.117, rely=0.511, height=26, width=67)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="black")
        self.Button2.configure(background="#919191")
        self.Button2.configure(disabledforeground="#adadad")
        self.Button2.configure(foreground="white")
        self.Button2.configure(highlightbackground="#919191")
        self.Button2.configure(highlightcolor="white")
        self.Button2.configure(text='''Execute''')

        # Progress bar
        self.progressbar = ttk.Progressbar(self.top, orient=tk.HORIZONTAL, length=200, mode='indeterminate')

    def connect_to_database(self):
        print("Clicked")  # Check if button click event is captured
        self.progressbar.place(relx=0.05, rely=0.2)  # Place the progress bar
        self.progressbar.start()  # Start the progress bar animation

        try:
            script_path = os.path.join(_location, "..", "Database", "create_database.py")
            print("Script Path:", script_path)  # Print out the script path for verification
            subprocess.run(["python", script_path])
        except Exception as e:
            print("Error:", e)  # Print out any exception messages for diagnosis

        self.progressbar.stop()  # Stop the progress bar animation
        self.progressbar.place_forget()  # Hide the progress bar

def main():
    root = tk.Tk()
    app = Toplevel1(root)  # Assuming Toplevel1 is the main class
    root.mainloop()

if __name__ == "__main__":
    main()
