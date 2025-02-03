from tkinter import *
import customtkinter

#customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
#customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# root = Tk()
root = customtkinter.CTk()

root.title('Tkinter.com - Custom Tkinter Switch')
# root.iconbitmap('images/codemy.ico')
root.geometry('700x300')

# Create Function
def switcher():
    if switch_var.get() == "Pre-Prod":
        my_switch.configure(text="Pre-Production")
    else:
        my_switch.configure(text="Production")

# Create Toggle function
def clicker():
    my_switch.toggle()

# Create a StringVar
switch_var = customtkinter.StringVar(value="Pre-Prod")

# Create Switch
my_switch = customtkinter.CTkSwitch(root, text="Pre-Production", command=switcher,
                                    variable=switch_var, onvalue="Pre-Prod", offvalue="Prod",
                                    switch_width=40,
                                    switch_height=16,
                                    font=("Helvetica", 24),
                                    text_color="blue",
                                    state="normal",
                                    )
my_switch.pack(pady=40)

root.mainloop()