import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk  # Import from Pillow
from FileConsolidation import file_consolidation  # Import the refactored function
#from dashfinal.main.penetration_code import filterandclean
import config  # Import your global config
from penetration_code import filterandclean


def browse_file():
    chosen_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if chosen_file:
        config.file_path = chosen_file.replace('\\', '/')  # Normalize and store in config
        user_name_start_index = config.file_path.lower().find('/users/') + 7
        user_name_end_index = config.file_path.find('/', user_name_start_index)
        if user_name_start_index > 6 and user_name_end_index > user_name_start_index:
            config.user_name = config.file_path[user_name_start_index:user_name_end_index]
            print(f"User Name Extracted: {config.user_name}")
        else:
            print("User folder not found in path")

        path_label.config(text=f"File Selected: {config.file_path}")
        browse_button.configure(style='Normal.TButton')
        root.update_idletasks()
    else:
        print("No file selected")

def process_file():
    if config.file_path:  # Check directly from config
        try:
            file_consolidation(config.file_path)  # Pass from config
            filterandclean()
            process_button.configure(style='Normal.TButton')
            root.update_idletasks()
            messagebox.showinfo("Success", "Processing complete. Check output in the file directory.")
        except Exception as e:
            process_button.configure(style='Error.TButton')
            root.update_idletasks()
            messagebox.showerror("Error", f"An error occurred: {e}")
    else:
        messagebox.showwarning("Warning", "Please select a file first!")


def on_enter(e):
    e.widget.configure(style='Hover.TButton')


def on_leave(e):
    e.widget.configure(style='TButton')


# Set up the GUI
root = tk.Tk()
root.title("File Consolidation Tool")
root.state('zoomed')  # Maximize window to full screen on start

# Set up the style
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('Normal.TButton', background='green', foreground='black')
style.configure('Hover.TButton', background='lightgrey', foreground='black')
style.configure('Error.TButton', background='red', foreground='black')

# Calculate the heights for the frames based on screen size
screen_height = root.winfo_screenheight()
top_frame_height = screen_height // 20  # Set to 1/3 of the screen height
bottom_frame_height = screen_height - top_frame_height  # Remaining for the bottom frame

# Create frames with adjusted heights and expanded filling
top_frame = tk.Frame(root, bg='black', height=top_frame_height)
top_frame.pack(fill='both', expand=True, side='top')

bottom_frame = tk.Frame(root, bg='#135C40', height=bottom_frame_height)
bottom_frame.pack(fill='both', expand=True)

# Load and display the BCG logo centered
logo_image = Image.open("bcgicon.png").convert("RGBA")
logo_image = logo_image.resize((120, 50), Image.Resampling.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = ttk.Label(top_frame, image=logo_photo, background='black')
logo_label.place(relx=0.5, rely=0.5, anchor='center')

# Add title label centered in the top_frame
title_label = ttk.Label(top_frame, text="Dashboard Data Consolidation Tool", font=('Helvetica', 24, 'bold'), background='black', foreground='white')
title_label.place(relx=0.5, rely=0.9, anchor='center')

# Configure buttons and labels
browse_button = ttk.Button(bottom_frame, text="Browse File", command=browse_file, style='TButton')
browse_button.pack(pady=(20, 10), anchor='n', padx=20)
browse_button.bind("<Enter>", on_enter)
browse_button.bind("<Leave>", on_leave)

path_label = ttk.Label(bottom_frame, text="No file selected", background='lightgrey', foreground='black', font=('Helvetica', 14))
path_label.pack(fill='x', padx=20, pady=(0, 10))

process_button = ttk.Button(bottom_frame, text="Process File", command=process_file, style='TButton')
process_button.pack(pady=(10, 20), anchor='s', padx=20)
process_button.bind("<Enter>", on_enter)
process_button.bind("<Leave>", on_leave)

# Start the GUI event loop
root.mainloop()
