import tkinter as tk
from tkinter import ttk, filedialog
import webbrowser
import os

class MilestonePreviewApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Milestone Preview")
        self.set_window_size_and_center(900, 600)  # Set the initial size and center the window

        # Create Notebook (Tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')

        # Add "Milestone Preview" tab
        self.preview_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.preview_tab, text="Milestone Preview")

        # Create a frame for the Text widget and scrollbars
        text_frame = ttk.Frame(self.preview_tab)
        text_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # HTML Input Textbox with Scrollbars
        self.html_input = tk.Text(text_frame, wrap='none')
        self.html_input.pack(side='left', fill='both', expand=True)

        # Vertical Scrollbar
        self.v_scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.html_input.yview)
        self.v_scrollbar.pack(side='right', fill='y')
        self.html_input['yscrollcommand'] = self.v_scrollbar.set

        # Horizontal Scrollbar
        self.h_scrollbar = ttk.Scrollbar(self.preview_tab, orient='horizontal', command=self.html_input.xview)
        self.h_scrollbar.pack(fill='x')
        self.html_input['xscrollcommand'] = self.h_scrollbar.set

        # Buttons Frame
        button_frame = ttk.Frame(self.preview_tab)
        button_frame.pack(fill='x', padx=5, pady=5)

        # Preview Button
        self.preview_button = ttk.Button(button_frame, text="Preview", command=self.preview_html)
        self.preview_button.pack(side='left', expand=True, padx=5)

        # Load File Button
        self.load_button = ttk.Button(button_frame, text="Load File", command=self.load_html_file)
        self.load_button.pack(side='left', expand=True, padx=5)

        # Save Milestone Button
        self.save_button = ttk.Button(button_frame, text="Save Milestone", command=self.save_milestone)
        self.save_button.pack(side='left', expand=True, padx=5)

        # Clear Button
        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_entry)
        self.clear_button.pack(side='right', expand=True, padx=5)

        # Variable to store HTML content
        self.html_string = ""

        # Bind right-click to show context menu
        self.html_input.bind("<Button-3>", self.show_context_menu)

        # Create context menu
        self.context_menu = tk.Menu(self.html_input, tearoff=0)
        self.context_menu.add_command(label="Copy", command=self.copy)
        self.context_menu.add_command(label="Paste", command=self.paste)

        # Bind keyboard shortcuts
        self.html_input.bind("<Control-c>", self.copy)
        self.html_input.bind("<Control-v>", self.paste)

    def set_window_size_and_center(self, width, height):
        """Set the window size and center it on the screen."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def preview_html(self):
        """Create a temporary HTML file and open it in the browser."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            return  # Do nothing if input is empty
        
        # Write to a temporary file
        temp_file = "temp_preview.html"
        with open(temp_file, "w", encoding="utf-8") as file:
            file.write(html_content)

        # Open in web browser
        webbrowser.open(f"file://{os.path.abspath(temp_file)}")

    def load_html_file(self):
        """Open an HTML or text file, clear the input box, and display its contents."""
        file_path = filedialog.askopenfilename(filetypes=[("HTML and Text Files", "*.html;*.htm;*.txt")])
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    self.html_string = file.read()
            except UnicodeDecodeError:
                with open(file_path, "r", encoding="latin-1") as file:
                    self.html_string = file.read()
            
            self.clear_entry()  # Clear the entry field before inserting new content
            self.html_input.insert("1.0", self.html_string)  # Insert file content

    def save_milestone(self):
        """Save the milestone content to a text file."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            return  # Do nothing if input is empty

        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(html_content)

    def clear_entry(self):
        """Clear the entry field."""
        self.html_input.delete("1.0", tk.END)

    def show_context_menu(self, event):
        """Show the context menu on right-click."""
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def copy(self, event=None):
        """Copy selected text to the clipboard."""
        self.html_input.event_generate("<<Copy>>")

    def paste(self, event=None):
        """Paste text from the clipboard."""
        self.html_input.event_generate("<<Paste>>")

if __name__ == "__main__":
    root = tk.Tk()
    app = MilestonePreviewApp(root)
    root.mainloop()