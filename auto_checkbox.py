import tkinter as tk
import tkinter.font as tkFont

# Assuming `root` is your Tkinter root window and `dynamic_vals` is a dictionary
root = tk.Tk()
dynamic_vals = {
    "112": {
        "cb_title": "Your Checkbox Title"
    }
}

# Initialize a dictionary to store the checkboxes
checkboxes = {}

# Number of checkboxes to create
num_checkboxes = 5

# Loop to create checkboxes
for cb_num in range(num_checkboxes):
    var = tk.IntVar(root)
    checkbox = tk.Checkbutton(root)
    checkbox["anchor"] = "w"
    
    # Set the font for the Checkbutton
    ft = tkFont.Font(family='Times', size=10)
    checkbox["font"] = ft
    
    # Set other properties for the Checkbutton
    checkbox["fg"] = "#333333"
    checkbox["justify"] = "left"
    checkbox["text"] = dynamic_vals["112"]["cb_title"]
    checkbox.place(x=10, y=196 + cb_num * 30, width=490, height=25)  # Adjust y position for each checkbox
    checkbox["offvalue"] = "0"
    checkbox["onvalue"] = "1"
    checkbox["variable"] = var
    
    # Select the Checkbutton by default
    checkbox.select()
    
    # Store the checkbox in the dictionary
    checkboxes[cb_num] = checkbox

# Start the Tkinter main loop
root.mainloop()