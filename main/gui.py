from tkinter import *
from tkinter import Toplevel, Label, Button, StringVar
from tkinter import ttk
import ctypes
from ctypes import wintypes
import os

def ask_student_action(parent, message):
    """
    Custom dialog to ask what action to take when a student already exists.
    """
    dialog = Toplevel(parent)
    dialog.title("Student Exists")
    dialog.transient(parent)
    dialog.grab_set()  # Makes the dialog modal

    # Set desired window dimensions
    window_width = 325
    window_height = 150  # Increase this value to make the dialog longer

    # Center the dialog on the screen
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x_position = int((screen_width / 2) - (window_width / 2))
    y_position = int((screen_height / 2) - (window_height / 2))
    dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    choice = StringVar()

    Label(dialog, text=message, wraplength=250).pack(pady=10)

    def set_choice(value):
        choice.set(value)
        dialog.destroy()

    Button(dialog, text="Modify", command=lambda: set_choice('modify')).pack(side='left', padx=10, pady=10)
    Button(dialog, text="Add Anyway", command=lambda: set_choice('add')).pack(side='left', padx=10, pady=10)
    Button(dialog, text="Cancel", command=lambda: set_choice('cancel')).pack(side='right', padx=10, pady=10)

    dialog.wait_window()  # Wait for the dialog to be closed
    return choice.get()

def center_window(root, width, height):
    # Get the screen dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate the available vertical space on Windows
    if os.name == 'nt':  # Check if the OS is Windows
        # Define a RECT structure for use with the Windows API
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", wintypes.LONG),
                ("top", wintypes.LONG),
                ("right", wintypes.LONG),
                ("bottom", wintypes.LONG)
            ]

        # Define constants for SystemParametersInfo function
        SPI_GETWORKAREA = 0x0030

        def get_taskbar_size():
            # Get screen dimensions
            screen_width = ctypes.windll.user32.GetSystemMetrics(0)
            screen_height = ctypes.windll.user32.GetSystemMetrics(1)

            # Get work area dimensions (excluding taskbar)
            work_area = RECT()
            ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(work_area), 0)

            # Calculate taskbar dimensions
            taskbar_width = screen_width - (work_area.right - work_area.left)
            taskbar_height = screen_height - (work_area.bottom - work_area.top)

            return taskbar_width, taskbar_height

        # Get the working area dimensions using the get_taskbar_size function
        taskbar_width, taskbar_height = get_taskbar_size()
        working_height = screen_height - taskbar_height - 30
        working_width = screen_width - taskbar_width
    else:
        # For macOS and other OS, use the total screen dimensions
        working_width = screen_width
        working_height = screen_height

    # Calculate position to center the window
    x = (working_width - width) // 2
    y = (working_height - height) // 2

    # Set the geometry of the window
    root.geometry(f'{width}x{height}+{x}+{y}')

def get_work_area():
    if os.name == 'nt':  # Windows
        # Define a RECT structure for use with the Windows API
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)
            ]

        SPI_GETWORKAREA = 0x0030
        work_area = RECT()
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(work_area), 0)

        screen_width = ctypes.windll.user32.GetSystemMetrics(0)
        screen_height = ctypes.windll.user32.GetSystemMetrics(1)

        working_width = work_area.right - work_area.left
        working_height = work_area.bottom - work_area.top

        return working_width, working_height

    elif platform.system() == 'Darwin':  # macOS
        from AppKit import NSScreen
        screen = NSScreen.mainScreen()
        screen_frame = screen.visibleFrame()
        working_width = int(screen_frame.size.width)
        working_height = int(screen_frame.size.height)
        return working_width, working_height

    else:
        # Default case for other platforms (Linux, etc.)
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        return screen_width, screen_height

def get_all_windows(root):
    windows = []
    def toplevels(window):
        for k, v in window.children.items():
            if isinstance(v, Toplevel):
                # print('Toplevel:', k, v)
                windows.append(v)    
            toplevels(v)
    toplevels(root)
    return windows

def get_all_treeviews(root):
    treeviews = []

    def find_treeviews(widget):
        # Recursively find all Treeview widgets
        for child in widget.winfo_children():
            if isinstance(child, ttk.Treeview):
                treeviews.append(child)
            else:
                find_treeviews(child)  # Recursively search through all child widgets

    find_treeviews(root)
    return treeviews

def calculate_row_height(font):
    return font.metrics("linespace") + 4

def update_all_treeviews(root, font):
    treeviews = get_all_treeviews(root)
    for treeview in treeviews:
        apply_treeview_style(treeview, font)  # Apply the style to each Treeview

def apply_treeview_style(treeview, font):
    style = ttk.Style(treeview)
    row_height = calculate_row_height(font)
    style.configure("Treeview", font=font, rowheight=row_height)
    style.configure("Treeview.Heading", font=font)

def adjust_window_size(window, scale_factor):
    # Get the current size of the window
    current_width = window.winfo_width()
    current_height = window.winfo_height()
    
    # Calculate the new size based on the scale factor
    new_width = int(current_width * scale_factor)
    new_height = int(current_height * scale_factor)
    
    # Get the working area dimensions (excluding taskbar/dock)
    working_width, working_height = get_work_area()
    
    # Ensure the new size does not exceed the screen's working area
    if new_width > working_width:
        new_width = working_width
    if new_height > working_height:
        new_height = working_height
    
    # Adjust the window size and center it
    # window.geometry(f"{new_width}x{new_height}")
    center_window(window, new_width, new_height)

def increase_window_sizes(root, scale_factor=1.1):
    # Iterate over all open windows
    for window in get_all_windows(root):
        if not window.state() == 'zoomed':  # Skip if the window is already maximized
            adjust_window_size(window, scale_factor)

def decrease_window_sizes(root, scale_factor=0.9):
    # Iterate over all open windows
    for window in get_all_windows(root):
        if not window.state() == 'zoomed':  # Skip if the window is already maximized
            adjust_window_size(window, scale_factor)

def right_click(event):
    # Function to print x and y coordinates when right-clicked
    x, y = event.x, event.y
    print(f"Right-clicked at ({x}, {y})")

def on_mouse_wheel(event, canvas):
    # Handle mouse wheel events for scrolling
    if event.delta:
        scroll_speed = int(event.delta / 120)
    else:
        scroll_speed = event.num if event.num != 0 else 0
    scroll_speed = -scroll_speed
    if scroll_speed == 0:
        scroll_speed = -1 if event.delta < 0 else 1
    canvas.yview_scroll(scroll_speed, "units")