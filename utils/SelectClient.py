import tkinter as tk
from tkinter import messagebox
import pandas as pd



def SelectClient():
    # Selects the client from the listbox and returns the order number
    print("Select Client from the listbox and enter the order number")

    # Load Client Data from Excel File and create a dataframe
    client_database = pd.read_excel("CellLineClients.xlsx")
    client_list = client_database["Nickname"].tolist()

    selected_item = ""
    order_number = ""

    def submit():
        # Get the selected client and order number
        nonlocal selected_item
        nonlocal order_number

        # Window Title "Please Select a Client"
        
        selected_item = listbox.get(listbox.curselection())
        order_number = order_entry.get()
        
        print("Selected Client:", selected_item)
        print("Order number:", order_number)

        # Close the window
        root.quit()
        root.destroy()

    root = tk.Tk()
    root.title("Order Form")
    root.geometry("300x300")

    # Create a listbox with the client names and a submit button
    label = tk.Label(root, text="Please Select a Client")


    listbox = tk.Listbox(root)
    for item in client_list:
        listbox.insert(tk.END, item)

    label.pack()
    listbox.pack()

    order_label = tk.Label(root, text="Order Number:")
    order_label.pack()

    order_entry = tk.Entry(root)
    order_entry.pack()

    submit_button = tk.Button(root, text="Submit", command=submit)
    submit_button.pack()

    root.mainloop()

    return selected_item, order_number

def display_readme():
    try:
        with open('readme.txt', 'r') as file:
            readme_content = file.read()
            messagebox.showinfo("Read Me", readme_content)
    except FileNotFoundError:
        messagebox.showerror("Error", "readme.txt not found.")

    # Create the main Tkinter window
    root = tk.Tk()

    # Set window title and size
    root.title("My Program")
    root.geometry("300x200")

    # Create a button to show the Read Me message
    readme_button = tk.Button(root, text="Read Me", command=display_readme)
    readme_button.pack(pady=50)

    # Start the Tkinter event loop
    root.mainloop()

def show_done_window():
    def show_done_message():
        messagebox.showinfo("Done", "Cell Line ID script has finished running!")
        root.destroy()  # Close the "Done" window and exit the program

    # Create the main Tkinter window
    root = tk.Tk()

    # Set window title and size
    root.title("Done")
    root.geometry("300x200")

    # Create a label to display the "Done" message
    message_label = tk.Label(root, text="Process completed!")
    message_label.pack(pady=50)

    # Create a button to close the window and exit the program
    done_button = tk.Button(root, text="Done", command=show_done_message)
    done_button.pack()

    # Start the Tkinter event loop
    root.mainloop()