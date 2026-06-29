import tkinter as tk
from tkinter import simpledialog, ttk, font, StringVar, IntVar, messagebox, Button
from tkcalendar import Calendar
from utils import *
from PIL import Image, ImageTk
if os.name == 'nt':
    import ctypes
    from ctypes import wintypes, windll
import os
import ctypes
import sqlite3
from pathlib import Path
import sys
from pathlib import Path

DB_FILE_PATH = PROJECT_ROOT / "database" / "master.db"
PROJECT_ROOT = Path(__file__).resolve().parent.parent

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scoring_data.tabe_scoring_tables import (
    tabe_scoring_tables,
    correct_answers,
)

# Edit tabe_scoring_tables and correct_answers in main.scoring_data with DRC scoring data

# If on Windows, set high DPI awareness
if os.name == 'nt':
    windll.shcore.SetProcessDpiAwareness(1)

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._completion_list = [item for item in kwargs.get("values", []) if isinstance(item, str)]
        self._hits = []
        self._typed_string = ""
        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<Tab>", self._accept_suggestion)

    def _on_key_release(self, event):
        """Triggered on each key release for autocomplete."""
        if event.keysym in ("BackSpace", "Left", "Right", "Up", "Down", "Tab"):
            return

        # Ignore Shift key and other modifiers
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return

        typed_string = self.get()
        if typed_string == "":
            self._reset()
            return

        # Match typed string with available options
        self._hits = [item for item in self._completion_list if item.lower().startswith(typed_string.lower())]

        # Display the closest match visually
        if self._hits:
            self._typed_string = typed_string
            self._show_suggestions()

    def _accept_suggestion(self, event):
        """Accept the suggestion when the Tab key is pressed."""
        if self._hits:
            self.set(self._hits[0])  # Set the combobox value to the first suggestion
            self.icursor(tk.END)  # Place cursor at the end
        return "break"  # Prevent default Tab key behavior

    def _show_suggestions(self):
        """Show the matched suggestions."""
        self.delete(0, tk.END)
        self.insert(0, self._hits[0])
        self.select_range(len(self._typed_string), tk.END)

    def _reset(self):
        """Reset the combobox suggestions."""
        self._hits = []
        self._typed_string = ""


def run_test_grading_gui(default_font, on_save=None):
    # Set up the main window and its components

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

    def get_student_data(DB_FILE_PATH):
        # Function to get student data from the database
        # Include Student.laces_id so the LACES field is populated from the Student table,
        # not from old Testing records.
        conn = sqlite3.connect(DB_FILE_PATH)
        cursor = conn.cursor()
        query = """
            SELECT
                Student.student_id,
                last_name || ', ' || first_name AS student_name,
                site_name,
                Student.laces_id
            FROM Student
            LEFT JOIN StudentSites ON Student.student_id = StudentSites.student_id
            LEFT JOIN Sites ON StudentSites.site_pk = Sites.site_pk
            GROUP BY Student.student_id, last_name, first_name
            ORDER BY last_name ASC
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        conn.close()
        return rows

    def center_window(answer_sheet_window, width, height):
        # Function to center a window on the screen
        screen_width = answer_sheet_window.winfo_screenwidth()
        screen_height = answer_sheet_window.winfo_screenheight()

        if os.name == 'nt':  # Windows-specific taskbar size adjustment
            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", wintypes.LONG),
                    ("top", wintypes.LONG),
                    ("right", wintypes.LONG),
                    ("bottom", wintypes.LONG)
                ]

            SPI_GETWORKAREA = 0x0030

            def get_taskbar_size():
                screen_width = ctypes.windll.user32.GetSystemMetrics(0)
                screen_height = ctypes.windll.user32.GetSystemMetrics(1)
                work_area = RECT()
                ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(work_area), 0)
                taskbar_width = screen_width - (work_area.right - work_area.left)
                taskbar_height = screen_height - (work_area.bottom - work_area.top)
                return taskbar_width, taskbar_height

            taskbar_width, taskbar_height = get_taskbar_size()
            working_height = screen_height - taskbar_height - 30
            working_width = screen_width - taskbar_width
        else:
            working_width = screen_width
            working_height = screen_height

        x = (working_width - width) // 2
        y = (working_height - height) // 2
        answer_sheet_window.geometry(f'{width}x{height}+{x}+{y}')

    def calculate_scaled_scores(level, form, reading_correct, writing_correct):
        # Hardcoded data from the Excel sheets for various levels and forms

        # Determine the correct sheet based on the level and form
        reading_sheet = f"Reading {level}{form}"
        writing_sheet = f"Writing {level}{form}"
        
        # Look up the reading score
        reading_data = tabe_scoring_tables[reading_sheet]
        writing_data = tabe_scoring_tables[writing_sheet]

        reading_index = reading_data['number_correct'].index(reading_correct)
        writing_index = writing_data['number_correct'].index(writing_correct)

        reading_scale_score = reading_data['scale_score'][reading_index]
        reading_nrs_value = reading_data['NRS'][reading_index]
        
        writing_scale_score = writing_data['scale_score'][writing_index]
        writing_nrs_value = writing_data['NRS'][writing_index]
        
        combined_scale_score = (reading_scale_score + writing_scale_score) // 2
        
        # result_message = (f"Student Name: {student_name}\n"
        #                   f"Level: {level}, Form: {form}\n"
        #                   f"Reading Scale Score: {reading_scale_score}, Reading NRS Value: {reading_nrs_value}\n"
        #                   f"Writing Scale Score: {writing_scale_score}, Writing NRS Value: {writing_nrs_value}\n"
        #                   f"Combined Scale Score: {combined_scale_score}\n"
        #                   f"{'-'*50}\n")
        # result = [reading_scale_score, reading_nrs_value, writing_scale_score, writing_nrs_value, combined_scale_score]
        result = {
            "Reading Raw": reading_correct,
            "Reading Scaled": reading_scale_score,
            "Reading NRS": reading_nrs_value,
            "Writing Raw": writing_correct,
            "Writing Scaled": writing_scale_score,
            "Writing NRS": writing_nrs_value,
            "Combined": combined_scale_score
        }
        return result

    class LevelTwoTest:
        def __init__(self, add_test_window):
            ORIGINAL_WIDTH = 1200  # Replace with the actual width of the original image
            ORIGINAL_HEIGHT = 1600  # Replace with the actual height of the original image
            self.add_test_window = add_test_window
            self.answer_sheet_window = None
            self.image_label = None
            
            self.selected_answers = {question: tk.StringVar() for question in range(1, 26)}
            self.selected_answers_secondary = {question: tk.StringVar() for question in range(1, 21)}  # Only 20 questions
            self.selected_answers_tertiary = {question: tk.StringVar() for question in range(1, 6)}  # Only 6 questions for writing score
            
            self.reading_total_string_var = tk.IntVar()
            self.writing_subtotal_string_var = tk.IntVar()
            self.expository_writing_subtotal_string_var = tk.IntVar()
            self.writing_total_string_var = tk.IntVar()

            self.primary_button_references = {q: {} for q in range(1, 26)}
            self.secondary_button_references = {q: {} for q in range(1, 21)}  # Only 20 questions
            self.tertiary_button_references = {q: {} for q in range(1, 6)}  # Only 5 questions
            
            if os.name == "nt":
                self.starting_width = 64
                self.starting_width_tertiary = 792
                self.starting_height = 212
                self.secondary_starting_height = 1105  # Different vertical starting height for secondary set & tertiary
                self.tertiary_starting_height = 1100
                
                self.height_offset = 47
                
                self.primary_width_offset = 168
                self.secondary_width_offset = 141
                self.tertiary_width_offset = 140
            else:
                self.starting_width = 61
                self.starting_height = 235
                self.starting_width_tertiary = 798
                self.secondary_starting_height = 1155  # Different vertical starting height for secondary set & tertiary
                
                self.height_offset = 47
                
                self.primary_width_offset = 169
                self.secondary_width_offset = 141
                self.tertiary_width_offset = 141

            self.original_image_size = (ORIGINAL_WIDTH, ORIGINAL_HEIGHT)  # Set these values according to your original image size
            if os.name != 'nt':
                self.button_size = (21, 21)
            else:
                self.button_size = (30, 30)  # Original size of the buttons

        def load_and_resize_image(self, new_size=None):
            # Constants for image path and original size
            IMAGE_PATH = os.path.join(get_test_png_path(), 'TABE Level 2 Cropped.png')
            MAX_SIZE = (800, 1000)  # Max width and height in pixels
            if os.name != 'nt':
                new_size = (1000, 1200)
            else:
                new_size = (1800, 1200)
            
            image = Image.open(IMAGE_PATH)

            # If a new size is provided, use it; otherwise, resize based on original dimensions
            if new_size:
                ratio = min(new_size[0] / image.width, new_size[1] / image.height)
            else:
                ratio = min(MAX_SIZE[0] / image.width, MAX_SIZE[1] / image.height)
            
            new_size = (int(image.width * ratio), int(image.height * ratio))
            resized_image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(resized_image), ratio, new_size


        def generate_question_positions(self, starting_width, starting_height, height_offset, primary_width_offset, questions_per_row=5, width_scale=1.0, height_scale=1.0):
            positions = {}
            current_width = starting_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 26):  # Assuming 25 questions
                positions[question] = {
                    'A': (current_width, current_height),
                    'B': (current_width + 28 * width_scale, current_height),
                    'C': (current_width + 56 * width_scale, current_height),
                    'D': (current_width + 84 * width_scale, current_height),
                    'F': (current_width, current_height),    # Use the same positions for 'F'-'J'
                    'G': (current_width + 28 * width_scale, current_height),
                    'H': (current_width + 56 * width_scale, current_height),
                    'J': (current_width + 84 * width_scale, current_height)
                }
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += primary_width_offset * width_scale
            
            return positions

        def generate_secondary_question_positions(self, starting_width, starting_height, height_offset, secondary_width_offset, questions_per_row=5, width_scale=1.0, height_scale=1.0):
            positions = {}
            current_width = starting_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 21):  # Assuming 20 questions in the secondary set
                positions[question] = {
                    'A': (current_width, current_height),
                    'B': (current_width + 28 * width_scale, current_height),
                    'C': (current_width + 56 * width_scale, current_height),
                    'F': (current_width, current_height),    # Use the same positions for 'F', 'G', 'H'
                    'G': (current_width + 28 * width_scale, current_height),
                    'H': (current_width + 56 * width_scale, current_height)
                }
                
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += secondary_width_offset * width_scale
            
            return positions
        
        def generate_tertiary_question_positions(
                self, 
                starting_width_tertiary, 
                starting_height, 
                height_offset, 
                tertiary_width_offset, 
                questions_per_row=5, 
                width_scale=1.0, 
                height_scale=1.0
            ):
            positions = {}
            current_width = starting_width_tertiary * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 6):
                positions[question] = {
                    '0': (current_width, current_height),
                    '1': (current_width + 28 * width_scale, current_height),
                    '2': (current_width + 56 * width_scale, current_height),
                    '3': (current_width + 84 * width_scale, current_height),
                    '4': (current_width + 112 * width_scale, current_height),
                }
                
                # Increment height for the next question
                current_height += (height_offset+2) * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += tertiary_width_offset * width_scale
            
            return positions

        def update_test_display(self, scaling_factor=1.0):
            # Resize the image based on the scaling factor
            new_image_size = (int(self.original_image_size[0] * scaling_factor), int(self.original_image_size[1] * scaling_factor))
            test_image, ratio, new_size = self.load_and_resize_image(new_size=new_image_size)
            self.image_label.config(image=test_image)
            self.image_label.image = test_image
            
            # Calculate scaling factors relative to the original image size
            width_scale = new_size[0] / self.original_image_size[0]
            height_scale = new_size[1] / self.original_image_size[1]

            # Resize buttons
            new_button_size = (int(self.button_size[0] * scaling_factor), int(self.button_size[1] * scaling_factor))

            # Generate question positions based on the scaled values
            question_positions = self.generate_question_positions(self.starting_width, self.starting_height, self.height_offset, self.primary_width_offset, width_scale=width_scale, height_scale=height_scale)
            secondary_question_positions = self.generate_secondary_question_positions(self.starting_width, self.secondary_starting_height, self.height_offset, self.secondary_width_offset, width_scale=width_scale, height_scale=height_scale)
            tertiary_question_positions = self.generate_tertiary_question_positions(self.starting_width_tertiary, self.tertiary_starting_height, self.height_offset, self.tertiary_width_offset, width_scale=width_scale, height_scale=height_scale)

            # Create primary set of radiobuttons
            for question in range(1, 26):
                question_var = self.selected_answers[question]
                
                # Alternate between "ABCD" and "FGHJ" for each question
                if question % 2 == 1:
                    answer_set = ['A', 'B', 'C', 'D']
                else:
                    answer_set = ['F', 'G', 'H', 'J']

                for answer in answer_set:
                    x, y = question_positions[question][answer]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="primary", button_size=new_button_size)

            # Create secondary set of radiobuttons with "ABC" for odd and "FGH" for even questions
            for question in range(1, 21):  # Only 20 questions in the secondary set
                question_var = self.selected_answers_secondary[question]
                
                # Alternate between "ABC" and "FGH" for each question
                if question % 2 == 1:
                    answer_set = ['A', 'B', 'C']
                else:
                    answer_set = ['F', 'G', 'H']

                for i, answer in enumerate(answer_set):
                    x, y = secondary_question_positions[question][chr(ord('A') + i)]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="secondary", button_size=new_button_size)

            # Create tertiary set of radiobuttons for writing score
            for question in range(1, 6):  # Only 5 questions in the tertiary set
                question_var = self.selected_answers_tertiary[question]
                
                # Alternate between "ABC" and "FGH" for each question
                if question != 5:
                    answer_set = ['0', '1', '2', '3']
                else:
                    answer_set = ['0', '1', '2', '3', '4']

                for i, answer in enumerate(answer_set):
                    x, y = tertiary_question_positions[question][str(0 + i)]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="tertiary", button_size=new_button_size)

            button_font = ("Segoe UI", 14)  # Font family, size

            if os.name == 'nt':
                grade_reading_button = tk.Button(self.content_frame, text="Grade", font=button_font, command=lambda: self.check_answers('reading'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_reading_button.place(x=1375, y=260)

                grade_writing_button = tk.Button(self.content_frame, text="Grade", font=button_font, command=lambda: self.check_answers('writing'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_writing_button.place(x=1377, y=1123) 
            else:
                grade_reading_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('reading'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_reading_button.place(x=910, y=176)

                grade_writing_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('writing'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_writing_button.place(x=910, y=752)

            finish_button = tk.Button(self.content_frame, font=button_font, text="Finish", command=self.get_scaled_scores, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
            self.content_frame.update_idletasks()
            
            # Position the button in the bottom-left corner
            if os.name == 'nt':
                finish_button.place(x=15, y=1123)
            else:
                finish_button.place(x=15, y=752)

        def create_round_radiobutton(self, question, variable, value, x, y, question_set="", button_size=(15, 15)):
            # Paths for normal and selected state images
            png_normal_path = os.path.join(get_test_png_path(), f"{value.lower()}_button.png")
            png_selected_path = os.path.join(get_test_png_path(), "filled_button.png")

            # Load and resize images
            normal_image = Image.open(png_normal_path).resize(button_size, Image.LANCZOS)
            selected_image = Image.open(png_selected_path).resize(button_size, Image.LANCZOS)
            
            normal_image = ImageTk.PhotoImage(normal_image)
            selected_image = ImageTk.PhotoImage(selected_image)
            
            def update_primary_buttons():
                for q in range(1, 26):
                    for ans in ['A', 'B', 'C', 'D', 'F', 'G', 'H', 'J']:
                        btn = self.primary_button_references[q].get(ans)
                        if btn and self.selected_answers[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)

            def update_secondary_buttons():
                for q in range(1, 21):
                    for ans in ['A', 'B', 'C', 'F', 'G', 'H']:  # Handle the secondary set
                        btn = self.secondary_button_references[q].get(ans)
                        if btn and self.selected_answers_secondary[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)

            def update_tertiary_buttons():
                for q in range(1, 6):
                    for ans in ['0', '1', '2', '3', '4']:  # Handle the tertiary (writing score) set
                        btn = self.tertiary_button_references[q].get(ans)
                        if btn and self.selected_answers_tertiary[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)
            
            # Create a Radiobutton with the normal image
            radiobutton = tk.Radiobutton(
                self.content_frame, 
                image=normal_image, 
                variable=variable, 
                value=value,
                indicatoron=False, 
                width=button_size[0], 
                height=button_size[1], 
                bd=0, 
                command=(update_primary_buttons if question_set=='primary' else update_secondary_buttons if question_set=='secondary' else update_tertiary_buttons))
            radiobutton.normal_image = normal_image  # Keep references to the images
            radiobutton.selected_image = selected_image
            radiobutton.place(x=x, y=y)
            
            # Store a reference to the radiobutton in the appropriate dictionary
            if question_set=='primary':
                self.primary_button_references[question][value] = radiobutton
            elif question_set=='secondary':
                self.secondary_button_references[question][value] = radiobutton
            else:
                self.tertiary_button_references[question][value] = radiobutton

        def enter_score(self):
            pass

        def save_test(self):
            # First, determine if student got MSG
            conn = sqlite3.connect(DB_FILE_PATH)
            cursor = conn.cursor()
            query = """
                SELECT
                    reading_nrs,
                    writing_nrs,
                    MAX(date)
                FROM
                    Testing
                WHERE
                    student_id = ?
            """
            cursor.execute(query, (fields_dict['Student ID'].get(),))
            last_test = cursor.fetchone() 
            conn.close()

            if last_test[0]:
                if '+' in str(self.scaled_scores["Reading NRS"]):
                    current_reading_nrs = int(str(self.scaled_scores["Reading NRS"]).replace('+', '')) + 1
                else:
                    current_reading_nrs = int(str(self.scaled_scores["Reading NRS"]))

                if '+' in str(self.scaled_scores["Writing NRS"]):
                    current_writing_nrs = int(str(self.scaled_scores["Writing NRS"]).replace('+', '')) + 1
                else:
                    current_writing_nrs = int(str(self.scaled_scores["Writing NRS"]))
                
                if '+' in str(last_test[0]):
                    last_reading_nrs = int(str(last_test[0]).replace('+', '')) + 1
                else:
                    last_reading_nrs = int(str(last_test[0]))

                if '+' in str(last_test[1]):
                    last_writing_nrs = int(str(last_test[1]).replace('+', ''))
                else:
                    last_writing_nrs = int(str(last_test[1]))

                msg_area = ""
                got_msg = 0
                if current_reading_nrs > last_reading_nrs:
                    got_msg = 1
                    msg_area += f'Reading: {last_reading_nrs} -> {current_reading_nrs}'
                if current_writing_nrs > last_writing_nrs:
                    got_msg = 1
                    if msg_area != "":
                        msg_area += f'\nWriting: {last_writing_nrs} -> {current_writing_nrs}'    
                    else:
                        msg_area += f'Writing: {last_writing_nrs} -> {current_writing_nrs}'
            else:
                got_msg = 0
                msg_area = ""

            # Get correct date string
            # formatted_date = format_date(fields_dict['Date:'].get().replace('/', '-'), out=True)
            # print(formatted_date)
            conn = sqlite3.connect(DB_FILE_PATH)
            cursor = conn.cursor()

            query = """
                INSERT INTO Testing (
                    student_id,
                    laces,
                    site,
                    staff,
                    primary_class,
                    date,
                    pre_or_post,
                    level,
                    form,
                    reading_raw,
                    reading_scaled,
                    reading_nrs,
                    writing_raw,
                    writing_scaled,
                    writing_nrs,
                    combined,
                    reading_answers,
                    writing_answers,
                    writing_folio_answers,
                    got_msg,
                    msg_area
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.execute(query, (
                fields_dict['Student ID'].get(),
                get_laces_id_for_save(),
                fields_dict['Site:'].get(),
                fields_dict['Aspire Staff:'].get(),
                fields_dict['Primary Class:'].get(),
                format_date(fields_dict['Date:'].get().replace('/', '-'), out=True),
                fields_dict['Pre/Post:'].get(),
                fields_dict['Level:'].get(),
                fields_dict['Form:'].get(),
                self.scaled_scores['Reading Raw'],
                self.scaled_scores['Reading Scaled'],
                self.scaled_scores['Reading NRS'],
                self.scaled_scores['Writing Raw'],
                self.scaled_scores['Writing Scaled'],
                self.scaled_scores['Writing NRS'],
                self.scaled_scores['Combined'],
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers.values()]),
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers_secondary.values()]),
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers_tertiary.values()]),
                got_msg,
                msg_area
            ))
            conn.commit()
            conn.close()

            # Repopulate TestingTab treeview
            if on_save:
                on_save()

        def check_answers(self, section):

            # Grade primary (reading) questions
            if section == 'reading':
                for i, (question, answer_choice) in enumerate(self.selected_answers.items()):
                    if i == 0:
                        self.reading_total_string_var.set(0)
                    answer_choice = answer_choice.get()
                    dict_index = f"{fields_dict[f'Level:'].get()}{fields_dict[f'Form:'].get()}"
                    correct_answer = correct_answers[dict_index]['reading'][i]
                    if answer_choice == correct_answer:
                        old_score = int(self.reading_total_string_var.get())
                        new_score = old_score + 1
                        self.reading_total_string_var.set(new_score)
            # Grade secondary (writing) questions
            else:
                for i, (question, answer_choice) in enumerate(self.selected_answers_secondary.items()):
                    if i == 0:
                        self.writing_subtotal_string_var.set(0)
                    answer_choice = answer_choice.get()
                    dict_index = f"{fields_dict[f'Level:'].get()}{fields_dict[f'Form:'].get()}"
                    correct_answer = correct_answers[dict_index]['writing'][i]
                    if answer_choice == correct_answer:
                        old_score = int(self.writing_subtotal_string_var.get())
                        new_score = old_score + 1
                        self.writing_subtotal_string_var.set(new_score)

                # Add writing folio score (tertiary questions)
                writing_folio_total = sum([int(v.get()) if v.get() else 0 for v in self.selected_answers_tertiary.values()])
                self.expository_writing_subtotal_string_var.set(writing_folio_total)

                # Add total writing folio score
                total_writing_score = int(self.writing_subtotal_string_var.get()) + int(self.expository_writing_subtotal_string_var.get())
                self.writing_total_string_var.set(total_writing_score)

            # Destroy existing labels before creating new ones
            if section == 'reading':
                if hasattr(self, 'grade_reading_label') and self.grade_reading_label.winfo_exists():
                    self.grade_reading_label.destroy()
            else:
                if hasattr(self, 'writing_subtotal_label') and self.writing_subtotal_label.winfo_exists():
                    self.writing_subtotal_label.destroy()
                if hasattr(self, 'expository_writing_subtotal_label') and self.expository_writing_subtotal_label.winfo_exists():
                    self.expository_writing_subtotal_label.destroy()
                if hasattr(self, 'writing_total_label') and self.writing_total_label.winfo_exists():
                    self.writing_total_label.destroy()

            # Display Reading Total label if score exists
            if section == 'reading' and self.reading_total_string_var.get():
                self.grade_reading_label = tk.Label(self.content_frame, textvariable=self.reading_total_string_var)
                if os.name == 'nt':
                    x_position = 1404 if int(self.reading_total_string_var.get()) < 10 else (1404 - 10)
                    self.grade_reading_label.place(x=x_position, y=183)
                    self.grade_reading_label.config(font=("Segoe UI", 22))
                else:
                    x_position = 929 if int(self.reading_total_string_var.get()) < 10 else 923
                    self.grade_reading_label.place(x=x_position, y=128)
                    self.grade_reading_label.config(font=("Segoe UI", 30))

            # Display Writing Total labels if scores exist
            if section == 'writing':
                if self.writing_subtotal_string_var.get():
                    self.writing_subtotal_label = tk.Label(self.content_frame, textvariable=self.writing_subtotal_string_var)
                    if os.name == 'nt':
                        x_position = 642 if int(self.writing_subtotal_string_var.get()) < 10 else (642 - 12)
                        self.writing_subtotal_label.place(x=x_position, y=1044)
                        self.writing_subtotal_label.config(font=("Segoe UI", 23))
                    else:
                        x_position = 423 if int(self.writing_subtotal_string_var.get()) < 10 else 417
                        self.writing_subtotal_label.place(x=x_position, y=702)
                        self.writing_subtotal_label.config(font=("Segoe UI", 30))

                if self.expository_writing_subtotal_string_var.get():
                    self.expository_writing_subtotal_label = tk.Label(self.content_frame, textvariable=self.expository_writing_subtotal_string_var)
                    if os.name == 'nt':
                        x_position = 1115 if int(self.expository_writing_subtotal_string_var.get()) < 10 else (1111 - 8)
                        self.expository_writing_subtotal_label.place(x=x_position, y=1044)
                        self.expository_writing_subtotal_label.config(font=("Segoe UI", 23))
                    else:
                        x_position = 740 if int(self.expository_writing_subtotal_string_var.get()) < 10 else 734
                        self.expository_writing_subtotal_label.place(x=x_position, y=702)
                        self.expository_writing_subtotal_label.config(font=("Segoe UI", 30))                    

                if self.writing_total_string_var.get():
                    self.writing_total_label = tk.Label(self.content_frame, textvariable=self.writing_total_string_var, anchor='center')
                    if os.name == 'nt':
                        x_position = 1404 if int(self.writing_total_string_var.get()) < 10 else (1404 - 10)
                        self.writing_total_label.place(x=x_position, y=1045)
                        self.writing_total_label.config(font=("Segoe UI", 23))
                    else:
                        x_position = 930 if int(self.writing_total_string_var.get()) < 10 else 924
                        self.writing_total_label.place(x=x_position, y=702)
                        self.writing_total_label.config(font=("Segoe UI", 30))
            
        def get_scaled_scores(self):
            scaled_scores = calculate_scaled_scores(
                fields_dict['Level:'].get(), 
                fields_dict['Form:'].get(),
                self.reading_total_string_var.get(),
                self.writing_total_string_var.get(),
                )
            
            self.scaled_scores = calculate_scaled_scores(
                fields_dict['Level:'].get(), 
                fields_dict['Form:'].get(),
                self.reading_total_string_var.get(),
                self.writing_total_string_var.get(),
                )
                
            demo_message = (
                f"Student: {fields_dict['Student Name:'].get()}\n"
                f"Test Level: {fields_dict['Level:'].get()}\n"
                f"Test Form: {fields_dict['Form:'].get()}\n"
                f"Date: {fields_dict['Date:'].get()}\n")
            score_message = "\n".join([f"{k}: {v}" for k,v in self.scaled_scores.items()])
            confirm_message = "\n\nDo you want to save the student's test data?"
        
            result_message = demo_message + score_message + confirm_message
            if messagebox.askyesno("Save Test Results?", result_message):
                self.save_test()
        
        def grade_test(self):
            self.answer_sheet_window = tk.Toplevel(self.add_test_window)
            self.answer_sheet_window.title((
                f'Level {fields_dict["Level:"].get()}, Form {fields_dict["Form:"].get()} - {fields_dict["Student Name:"].get()} - {fields_dict["Date:"].get()}')
            )
            
            # Setup canvas and scrollbars
            canvas = tk.Canvas(self.answer_sheet_window, width=1511, height=800)
            v_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.VERTICAL, command=canvas.yview)
            h_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.HORIZONTAL, command=canvas.xview)
            
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create a frame inside the canvas to hold all content
            self.content_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
            
            # Ensure the scroll region matches the content size
            self.content_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))

            # Bind the right-click event (Button-3 in Tkinter) to the function
            self.answer_sheet_window.bind("<Button-3>", right_click)
            canvas.bind_all("<MouseWheel>", lambda e: on_mouse_wheel(e, canvas))

            center_window(self.answer_sheet_window, 1600, 850)

            # Label to display the test image
            self.image_label = tk.Label(self.content_frame)
            self.image_label.pack()

            center_window(self.answer_sheet_window, 1530, 900)

            # Update the display with the constant image and size
            self.update_test_display()


    class LevelOneTest:
        def __init__(self, add_test_window):
            ORIGINAL_WIDTH = 1200  # Replace with the actual width of the original image
            ORIGINAL_HEIGHT = 1600  # Replace with the actual height of the original image
            self.add_test_window = add_test_window
            self.answer_sheet_window = None
            self.image_label = None
            
            self.selected_answers = {question: tk.StringVar() for question in range(1, 26)}
            self.selected_answers_secondary = {question: tk.StringVar() for question in range(1, 21)}  # Only 20 questions
            self.selected_answers_tertiary = {question: tk.StringVar() for question in range(1, 6)}  # Only 6 questions for writing score
            
            self.reading_total_string_var = tk.IntVar()
            self.writing_subtotal_string_var = tk.IntVar()
            self.expository_writing_subtotal_string_var = tk.IntVar()
            self.writing_total_string_var = tk.IntVar()

            self.primary_button_references = {q: {} for q in range(1, 26)}
            self.secondary_button_references = {q: {} for q in range(1, 21)}  # Only 20 questions
            self.tertiary_button_references = {q: {} for q in range(1, 6)}  # Only 5 questions

            
            if os.name == 'nt':
                self.starting_height = 177
                self.starting_width = 65
                self.starting_tertiary_width = 793
                self.secondary_starting_height = 1098
                self.height_offset = 48

                self.primary_width_offset = 139.5
                self.secondary_width_offset = 139.5
                self.tertiary_width_offset = 139.5
            else:
                self.starting_height = 195
                self.starting_width = 70
                self.secondary_starting_height = 1150  # Different vertical starting height for secondary set & tertiary
                self.height_offset = 52
            
                self.primary_width_offset = 138
                self.secondary_width_offset = 138
                self.tertiary_width_offset = 138

            self.original_image_size = (ORIGINAL_WIDTH, ORIGINAL_HEIGHT)  # Set these values according to your original image size
            if os.name != 'nt':
                self.button_size = (21, 21)
            else:
                self.button_size = (30, 30)  # Original size of the buttons


        def load_and_resize_image(self, new_size=None):
            # IMAGE_PATH = os.path.join(get_test_png_path(), 'TABE Class E Level 1_page_1.png')
            IMAGE_PATH = os.path.join(get_test_png_path(), 'TABE Level 1 Cropped.png')
            image = Image.open(IMAGE_PATH)
            MAX_SIZE = (800, 1000)  # Max width and height in pixels
            if os.name != 'nt':
                new_size = (1000, 1200)
            else:
                new_size = (1800, 1200)

            # If a new size is provided, use it; otherwise, resize based on original dimensions
            if new_size:
                ratio = min(new_size[0] / image.width, new_size[1] / image.height)
            else:
                ratio = min(MAX_SIZE[0] / image.width, MAX_SIZE[1] / image.height)
            
            new_size = (int(image.width * ratio), int(image.height * ratio))
            resized_image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(resized_image), ratio, new_size


        def generate_question_positions(
                self, starting_width, 
                starting_height, 
                height_offset, 
                primary_width_offset,
                questions_per_row=5, 
                width_scale=1.0, 
                height_scale=1.0
            ):
            positions = {}
            current_width = starting_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 26):  # Assuming 25 questions
                positions[question] = {
                    'A': (current_width, current_height),
                    'B': (current_width + 29 * width_scale, current_height),
                    'C': (current_width + 58 * width_scale, current_height),
                    'F': (current_width, current_height),    # Use the same positions for 'F'-'J'
                    'G': (current_width + 29 * width_scale, current_height),
                    'H': (current_width + 58 * width_scale, current_height),
                }
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += primary_width_offset * width_scale
            
            return positions

        def generate_secondary_question_positions(
                self, 
                starting_width, 
                starting_height, 
                height_offset, 
                secondary_width_offset, 
                questions_per_row=5, 
                width_scale=1.0, 
                height_scale=1.0
            ):
            positions = {}
            current_width = starting_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 21):  # Assuming 20 questions in the secondary set
                positions[question] = {
                    'A': (current_width, current_height),
                    'B': (current_width + 28 * width_scale, current_height),
                    'C': (current_width + 56 * width_scale, current_height),
                    'F': (current_width, current_height),    # Use the same positions for 'F', 'G', 'H'
                    'G': (current_width + 28 * width_scale, current_height),
                    'H': (current_width + 56 * width_scale, current_height)
                }
                
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += secondary_width_offset * width_scale
            
            return positions
        
        def generate_tertiary_question_positions(
                self, 
                tertiary_width, 
                starting_height, 
                height_offset, 
                tertiary_width_offset, 
                questions_per_row=5, 
                width_scale=1.0, 
                height_scale=1.0
            ):
            positions = {}
            current_width = tertiary_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 6):
                positions[question] = {
                    '0': (current_width, current_height),
                    '1': (current_width + 28 * width_scale, current_height),
                    '2': (current_width + 56 * width_scale, current_height),
                    '3': (current_width + 84 * width_scale, current_height),
                    '4': (current_width + 112 * width_scale, current_height),
                }
                
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += tertiary_width_offset * width_scale
            
            return positions

        def update_test_display(self, scaling_factor=1.0):
            # Resize the image based on the scaling factor
            new_image_size = (int(self.original_image_size[0] * scaling_factor), int(self.original_image_size[1] * scaling_factor))
            test_image, ratio, new_size = self.load_and_resize_image(new_size=new_image_size)
            self.image_label.config(image=test_image)
            self.image_label.image = test_image
            
            # Calculate scaling factors relative to the original image size
            width_scale = new_size[0] / self.original_image_size[0]
            height_scale = new_size[1] / self.original_image_size[1]

            # Resize buttons
            new_button_size = (int(self.button_size[0] * scaling_factor), int(self.button_size[1] * scaling_factor))

            # Generate question positions based on the scaled values
            question_positions = self.generate_question_positions(self.starting_width, self.starting_height, self.height_offset, self.primary_width_offset, width_scale=width_scale, height_scale=height_scale)
            secondary_question_positions = self.generate_secondary_question_positions(self.starting_width, self.secondary_starting_height, self.height_offset, self.secondary_width_offset, width_scale=width_scale, height_scale=height_scale)
            tertiary_question_positions = self.generate_tertiary_question_positions(self.starting_tertiary_width, self.secondary_starting_height, self.height_offset, self.tertiary_width_offset, width_scale=width_scale, height_scale=height_scale)

            # Create primary set of radiobuttons
            for question in range(1, 26):
                question_var = self.selected_answers[question]
                
                # Alternate between "ABCD" and "FGHJ" for each question
                if question % 2 == 1:
                    answer_set = ['A', 'B', 'C']
                else:
                    answer_set = ['F', 'G', 'H']

                for answer in answer_set:
                    x, y = question_positions[question][answer]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="primary", button_size=new_button_size)

            # Create secondary set of radiobuttons with "ABC" for odd and "FGH" for even questions
            for question in range(1, 21):  # Only 20 questions in the secondary set
                question_var = self.selected_answers_secondary[question]
                
                # Alternate between "ABC" and "FGH" for each question
                if question % 2 == 1:
                    answer_set = ['A', 'B', 'C']
                else:
                    answer_set = ['F', 'G', 'H']

                for i, answer in enumerate(answer_set):
                    x, y = secondary_question_positions[question][chr(ord('A') + i)]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="secondary", button_size=new_button_size)

            # Create tertiary set of radiobuttons for writing score
            for question in range(1, 6):  # Only 5 questions in the tertiary set
                question_var = self.selected_answers_tertiary[question]
                
                # Alternate between "ABC" and "FGH" for each question
                if question != 5:
                    answer_set = ['0', '1', '2', '3']
                else:
                    answer_set = ['0', '1', '2', '3', '4']

                for i, answer in enumerate(answer_set):
                    # x, y = tertiary_question_positions[question][chr(ord('1') + i)]
                    x, y = tertiary_question_positions[question][str(0 + i)]
                    self.create_round_radiobutton(question, question_var, answer, x, y, question_set="tertiary", button_size=new_button_size)
        
            button_font = ("Segoe UI", 14)  # Font family, size

            if os.name == 'nt':
                grade_reading_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('reading'), font=button_font, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_reading_button.place(x=1427, y=237)

                grade_writing_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('writing'), font=button_font, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_writing_button.place(x=1427, y=1130)
            else:
                grade_reading_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('reading'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_reading_button.place(x=901, y=155)

                grade_writing_button = tk.Button(self.content_frame, text="Grade",command=lambda: self.check_answers('writing'), borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_writing_button.place(x=902, y=717) 

            finish_button = tk.Button(self.content_frame, font=button_font, text="Finish", command=self.get_scaled_scores, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
            self.content_frame.update_idletasks()
            
            # Position the button in the bottom-left corner
            if os.name == 'nt':
                finish_button.place(x=15, y=1130)
            else:
                finish_button.place(x=15, y=717)
            
            # finish_button.place(x=10, y=200)  # Adjust `y` as needed
            # self.answer_sheet_window.update_idletasks()
            # finish_button.place(x=10, y=self.answer_sheet_window.winfo_height() - finish_button.winfo_height() - 10)
        
        def create_round_radiobutton(self, question, variable, value, x, y, question_set="", button_size=(15, 15)):
            # Paths for normal and selected state images
            png_normal_path = os.path.join(get_test_png_path(), f"{value.lower()}_button_lv_1.png")
            png_selected_path = os.path.join(get_test_png_path(), f"filled_lvl_1_button.png")

            # Load and resize images
            normal_image = Image.open(png_normal_path).resize(button_size, Image.LANCZOS)
            selected_image = Image.open(png_selected_path).resize(button_size, Image.LANCZOS)
            
            normal_image = ImageTk.PhotoImage(normal_image)
            selected_image = ImageTk.PhotoImage(selected_image)
            
            def update_primary_buttons():
                for q in range(1, 26):
                    for ans in ['A', 'B', 'C', 'F', 'G', 'H']:
                        btn = self.primary_button_references[q].get(ans)
                        if btn and self.selected_answers[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)

            def update_secondary_buttons():
                for q in range(1, 21):
                    for ans in ['A', 'B', 'C', 'F', 'G', 'H']:  # Handle the secondary set
                        btn = self.secondary_button_references[q].get(ans)
                        if btn and self.selected_answers_secondary[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)

            def update_tertiary_buttons():
                for q in range(1, 6):
                    for ans in ['0', '1', '2', '3', '4']:  # Handle the tertiary (writing score) set
                        btn = self.tertiary_button_references[q].get(ans)
                        if btn and self.selected_answers_tertiary[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)
            
            # Create a Radiobutton with the normal image
            radiobutton = tk.Radiobutton(
                self.content_frame, 
                image=normal_image, 
                variable=variable, 
                value=value,
                indicatoron=False, 
                width=button_size[0], 
                height=button_size[1], 
                bd=0, 
                command=(update_primary_buttons if question_set=='primary' else update_secondary_buttons if question_set=='secondary' else update_tertiary_buttons))
            radiobutton.normal_image = normal_image  # Keep references to the images
            radiobutton.selected_image = selected_image
            radiobutton.place(x=x, y=y)
            
            # Store a reference to the radiobutton in the appropriate dictionary
            if question_set=='primary':
                self.primary_button_references[question][value] = radiobutton
            elif question_set=='secondary':
                self.secondary_button_references[question][value] = radiobutton
            else:
                self.tertiary_button_references[question][value] = radiobutton

        def enter_score(self):
            pass

        def save_test(self):
            # First, determine if student got MSG
            conn = sqlite3.connect(DB_FILE_PATH)
            cursor = conn.cursor()
            query = """
                SELECT
                    reading_nrs,
                    writing_nrs,
                    MAX(date)
                FROM
                    Testing
                WHERE
                    student_id = ?
            """
            cursor.execute(query, (fields_dict['Student ID'].get(),))
            last_test = cursor.fetchone()
            conn.close()

            # if last_test[0]:
            #     current_reading_nrs = int(str(self.scaled_scores["Reading NRS"]).replace('+', ''))
            #     current_writing_nrs = int(str(self.scaled_scores["Writing NRS"]).replace('+', ''))
            #     last_reading_nrs = int(str(last_test[0]).replace('+', ''))
            #     last_writing_nrs = int(str(last_test[1]).replace('+', ''))

            if last_test[0]:
                if '+' in str(self.scaled_scores["Reading NRS"]):
                    current_reading_nrs = int(str(self.scaled_scores["Reading NRS"]).replace('+', '')) + 1
                else:
                    current_reading_nrs = int(str(self.scaled_scores["Reading NRS"]))

                if '+' in str(self.scaled_scores["Writing NRS"]):
                    current_writing_nrs = int(str(self.scaled_scores["Writing NRS"]).replace('+', '')) + 1
                else:
                    current_writing_nrs = int(str(self.scaled_scores["Writing NRS"]))
                
                if '+' in str(last_test[0]):
                    last_reading_nrs = int(str(last_test[0]).replace('+', '')) + 1
                else:
                    last_reading_nrs = int(str(last_test[0]))

                if '+' in str(last_test[1]):
                    last_writing_nrs = int(str(last_test[1]).replace('+', ''))
                else:
                    last_writing_nrs = int(str(last_test[1]))


                msg_area = ""
                got_msg = 0
                if current_reading_nrs > last_reading_nrs:
                    got_msg = 1
                    msg_area += f'Reading: {last_reading_nrs} -> {current_reading_nrs}'
                if current_writing_nrs > last_writing_nrs:
                    got_msg = 1
                    if msg_area != "":
                        msg_area += f'\nWriting: {last_writing_nrs} -> {current_writing_nrs}'    
                    else:
                        msg_area += f'Writing: {last_writing_nrs} -> {current_writing_nrs}'
            else:
                got_msg = 0
                msg_area = ""

            # Get correct date string
            # formatted_date = format_date(fields_dict['Date:'].get().replace('/', '-'), out=True)
            # print(formatted_date)
            conn = sqlite3.connect(DB_FILE_PATH)
            cursor = conn.cursor()

            query = """
                INSERT INTO Testing (
                    student_id,
                    laces,
                    site,
                    staff,
                    primary_class,
                    date,
                    pre_or_post,
                    level,
                    form,
                    reading_raw,
                    reading_scaled,
                    reading_nrs,
                    writing_raw,
                    writing_scaled,
                    writing_nrs,
                    combined,
                    reading_answers,
                    writing_answers,
                    writing_folio_answers,
                    got_msg,
                    msg_area
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.execute(query, (
                fields_dict['Student ID'].get(),
                get_laces_id_for_save(),
                fields_dict['Site:'].get(),
                fields_dict['Aspire Staff:'].get(),
                fields_dict['Primary Class:'].get(),
                format_date(fields_dict['Date:'].get().replace('/', '-'), out=True),
                fields_dict['Pre/Post:'].get(),
                fields_dict['Level:'].get(),
                fields_dict['Form:'].get(),
                self.scaled_scores['Reading Raw'],
                self.scaled_scores['Reading Scaled'],
                self.scaled_scores['Reading NRS'],
                self.scaled_scores['Writing Raw'],
                self.scaled_scores['Writing Scaled'],
                self.scaled_scores['Writing NRS'],
                self.scaled_scores['Combined'],
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers.values()]),
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers_secondary.values()]),
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers_tertiary.values()]),
                got_msg,
                msg_area
            ))
            conn.commit()
            conn.close()

            if on_save:
                on_save()

        def get_scaled_scores(self):
            self.scaled_scores = calculate_scaled_scores(
                fields_dict['Level:'].get(), 
                fields_dict['Form:'].get(),
                self.reading_total_string_var.get(),
                self.writing_total_string_var.get(),
                )
                
            demo_message = (
                f"Student: {fields_dict['Student Name:'].get()}\n"
                f"Test Level: {fields_dict['Level:'].get()}\n"
                f"Test Form: {fields_dict['Form:'].get()}\n"
                f"Date: {fields_dict['Date:'].get()}\n")
            score_message = "\n".join([f"{k}: {v}" for k,v in self.scaled_scores.items()])
            confirm_message = "\n\nDo you want to save the student's test data?"
        
            result_message = demo_message + score_message + confirm_message
            if messagebox.askyesno("Save Test Results?", result_message):
                self.save_test()
        
        def check_answers(self, section):

        # Grade primary (reading) questions
            if section == 'reading':
                self.reading_total_string_var.set(0)  # Initialize the score
                for i, (question, answer_choice) in enumerate(self.selected_answers.items()):
                    answer_choice = answer_choice.get()
                    dict_index = f"{fields_dict[f'Level:'].get()}{fields_dict[f'Form:'].get()}"
                    correct_answer = correct_answers[dict_index]['reading'][i]
                    if answer_choice == correct_answer:
                        new_score = int(self.reading_total_string_var.get()) + 1
                        self.reading_total_string_var.set(new_score)

            # Grade secondary (writing) questions
            else:
                self.writing_subtotal_string_var.set(0)  # Initialize the score
                for i, (question, answer_choice) in enumerate(self.selected_answers_secondary.items()):
                    answer_choice = answer_choice.get()
                    dict_index = f"{fields_dict[f'Level:'].get()}{fields_dict[f'Form:'].get()}"
                    correct_answer = correct_answers[dict_index]['writing'][i]
                    if answer_choice == correct_answer:
                        new_score = int(self.writing_subtotal_string_var.get()) + 1
                        self.writing_subtotal_string_var.set(new_score)

                # Add writing folio score (tertiary questions)
                writing_folio_total = sum([int(v.get()) if v.get() else 0 for v in self.selected_answers_tertiary.values()])
                self.expository_writing_subtotal_string_var.set(writing_folio_total)

                # Add total writing folio score
                total_writing_score = int(self.writing_subtotal_string_var.get()) + int(self.expository_writing_subtotal_string_var.get())
                self.writing_total_string_var.set(total_writing_score)

            # Destroy existing labels before creating new ones
            if section == 'reading':
                if hasattr(self, 'grade_reading_label') and self.grade_reading_label.winfo_exists():
                    self.grade_reading_label.destroy()
            else:
                if hasattr(self, 'writing_subtotal_label') and self.writing_subtotal_label.winfo_exists():
                    self.writing_subtotal_label.destroy()
                if hasattr(self, 'expository_writing_subtotal_label') and self.expository_writing_subtotal_label.winfo_exists():
                    self.expository_writing_subtotal_label.destroy()
                if hasattr(self, 'writing_total_label') and self.writing_total_label.winfo_exists():
                    self.writing_total_label.destroy()

            # Display Reading Total label if score exists
            if section == 'reading' and self.reading_total_string_var.get():
                self.grade_reading_label = tk.Label(self.content_frame, textvariable=self.reading_total_string_var)
                if os.name == 'nt':
                    x_position = 1453 if int(self.reading_total_string_var.get()) < 10 else (1453 - 10)
                    self.grade_reading_label.place(x=x_position, y=153)
                    self.grade_reading_label.config(font=("Segoe UI", 23))
                else:    
                    x_position = 923 if int(self.reading_total_string_var.get()) < 10 else 917
                    self.grade_reading_label.place(x=x_position, y=106)
                    self.grade_reading_label.config(font=("Segoe UI", 30))

            # Display Writing Total labels if scores exist
            if section == 'writing':
                if self.writing_subtotal_string_var.get():
                    self.writing_subtotal_label = tk.Label(self.content_frame, textvariable=self.writing_subtotal_string_var)
                    if os.name == 'nt':
                        x_position = 666 if int(self.writing_subtotal_string_var.get()) < 10 else (666 - 10)
                        self.writing_subtotal_label.place(x=x_position, y=1045)
                        self.writing_subtotal_label.config(font=("Segoe UI", 23))
                    else:    
                        x_position = 425 if int(self.writing_subtotal_string_var.get()) < 10 else 419 
                        self.writing_subtotal_label.place(x=x_position, y=670)
                        self.writing_subtotal_label.config(font=("Segoe UI", 30))

                if self.expository_writing_subtotal_string_var.get():
                    if os.name == 'nt':
                        expository_x_position = 1157 if int(self.expository_writing_subtotal_string_var.get()) < 10 else (1157 - 13)
                        self.expository_writing_subtotal_label = tk.Label(self.content_frame, textvariable=self.expository_writing_subtotal_string_var)
                        self.expository_writing_subtotal_label.place(x=expository_x_position, y=1045)
                        self.expository_writing_subtotal_label.config(font=("Segoe UI", 24))
                    else:
                        expository_x_position = 735 if int(self.expository_writing_subtotal_string_var.get()) < 10 else 730
                        self.expository_writing_subtotal_label = tk.Label(self.content_frame, textvariable=self.expository_writing_subtotal_string_var)
                        self.expository_writing_subtotal_label.place(x=expository_x_position, y=670)
                        self.expository_writing_subtotal_label.config(font=("Segoe UI", 30))

                if self.writing_total_string_var.get():
                    if os.name == 'nt':
                        writing_x_position = 1454 if int(self.writing_total_string_var.get()) < 10 else (1454 - 10)
                        self.writing_total_label = tk.Label(self.content_frame, textvariable=self.writing_total_string_var)
                        self.writing_total_label.place(x=writing_x_position, y=1045)
                        self.writing_total_label.config(font=("Segoe UI", 24))
                    else:
                        writing_x_position = 922 if int(self.writing_total_string_var.get()) < 10 else 916
                        self.writing_total_label = tk.Label(self.content_frame, textvariable=self.writing_total_string_var)
                        self.writing_total_label.place(x=writing_x_position, y=670)
                        self.writing_total_label.config(font=("Segoe UI", 30))

        def grade_test(self):
            self.answer_sheet_window = tk.Toplevel(self.add_test_window)
            self.answer_sheet_window.title((
                f'Level {fields_dict["Level:"].get()}, Form {fields_dict["Form:"].get()} - {fields_dict["Student Name:"].get()} - {fields_dict["Date:"].get()}')
            )
            
            # Setup canvas and scrollbars
            canvas = tk.Canvas(self.answer_sheet_window, width=1560, height=800)
            v_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.VERTICAL, command=canvas.yview)
            h_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.HORIZONTAL, command=canvas.xview)
            
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create a frame inside the canvas to hold all content
            self.content_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
            
            # Ensure the scroll region matches the content size
            self.content_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))

            # Bind the right-click event (Button-3 in Tkinter) to the function
            self.answer_sheet_window.bind("<Button-3>", right_click)
            canvas.bind_all("<MouseWheel>", lambda e: on_mouse_wheel(e, canvas))

            center_window(self.answer_sheet_window, 1600, 850)

            # Label to display the test image
            self.image_label = tk.Label(self.content_frame)
            self.image_label.pack()

            # Update the display with the constant image and size
            self.update_test_display()


    class LocatorTest:
        def __init__(self, add_test_window):
            ORIGINAL_WIDTH = 1200  # Replace with the actual width of the original image
            ORIGINAL_HEIGHT = 1600  # Replace with the actual height of the original image
            self.add_test_window = add_test_window
            self.answer_sheet_window = None
            self.image_label = None
            
            self.total_correct_answers_string_var = tk.IntVar()
            
            self.selected_answers = {question: tk.StringVar() for question in range(1, 16)}
            self.primary_button_references = {q: {} for q in range(1, 16)}
            
            if os.name == 'nt':
                self.starting_width = 149
                self.starting_height = 718
                self.height_offset = 63
                self.primary_width_offset = 138.5
                self.original_image_size = (ORIGINAL_WIDTH, ORIGINAL_HEIGHT)  # Set these values according to your original image size
            else:
                self.starting_width = 149
                self.starting_height = 720
                self.height_offset = 63
                self.primary_width_offset = 138.5
                self.original_image_size = (ORIGINAL_WIDTH, ORIGINAL_HEIGHT)  # Set these values according to your original image size
            if os.name == 'nt':
                self.button_size = (25, 25)  # Original size of the buttons
            else:
                self.button_size = (14, 14)


        def load_and_resize_image(self, new_size=None):
            IMAGE_PATH = os.path.join(get_test_png_path(), 'Locator Test Answer Key_page_1.png')
            image = Image.open(IMAGE_PATH)
            MAX_SIZE = (800, 1000)  # Max width and height in pixels
            if os.name != 'nt':
                new_size = (1000, 1200)

            # If a new size is provided, use it; otherwise, resize based on original dimensions
            if new_size:
                ratio = min(new_size[0] / image.width, new_size[1] / image.height)
            else:
                ratio = min(MAX_SIZE[0] / image.width, MAX_SIZE[1] / image.height)
            
            new_size = (int(image.width * ratio), int(image.height * ratio))
            resized_image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(resized_image), ratio, new_size


        def generate_question_positions(self, starting_width, starting_height, height_offset, primary_width_offset, questions_per_row=3, width_scale=1.0, height_scale=1.0):
            positions = {}
            current_width = starting_width * width_scale
            current_height = starting_height * height_scale

            for question in range(1, 16): 
                positions[question] = {
                    'A': (current_width + 1, current_height),
                    'B': (current_width + 24 * width_scale, current_height),
                    'C': (current_width + 47 * width_scale, current_height),
                    'D': (current_width + 70 * width_scale, current_height),
                    'F': (current_width + 1, current_height),    # Use the same positions for 'F'-'J'
                    'G': (current_width + 24 * width_scale, current_height),
                    'H': (current_width + 47 * width_scale, current_height),
                    'J': (current_width + 70 * width_scale, current_height),
                }
                # Increment height for the next question
                current_height += height_offset * height_scale
                
                # After `questions_per_row` questions, reset height and increase width
                if question % questions_per_row == 0:
                    current_height = starting_height * height_scale
                    current_width += primary_width_offset * width_scale
            
            return positions

        def update_test_display(self, scaling_factor=1.6):
            # Resize the image based on the scaling factor
            new_image_size = (int(self.original_image_size[0] * scaling_factor), int(self.original_image_size[1] * scaling_factor))
            test_image, ratio, new_size = self.load_and_resize_image(new_size=new_image_size)
            self.image_label.config(image=test_image)
            self.image_label.image = test_image
            
            # Calculate scaling factors relative to the original image size
            width_scale = new_size[0] / self.original_image_size[0]
            height_scale = new_size[1] / self.original_image_size[1]

            # Resize buttons
            new_button_size = (int(self.button_size[0] * (scaling_factor-.25)), int(self.button_size[1] * (scaling_factor-.25)))

            # Generate question positions based on the scaled values
            question_positions = self.generate_question_positions(self.starting_width, self.starting_height, self.height_offset, self.primary_width_offset, width_scale=width_scale, height_scale=height_scale)
            # Create primary set of radiobuttons
            for question in range(1, 16):
                question_var = self.selected_answers[question]
                
                # Alternate between "ABCD" and "FGHJ" for each question
                if question in [1, 3, 11, 13, 15]:
                    answer_set = ['A', 'B', 'C']
                elif question in [2, 10, 12, 14]:
                    answer_set = ['F', 'G', 'H']
                elif question in [5, 7, 9]:
                    answer_set = ['A', 'B', 'C', 'D']
                else:
                    answer_set = ['F', 'G', 'H', 'J']

                for answer in answer_set:
                    x, y = question_positions[question][answer]
                    self.create_round_radiobutton(question, question_var, answer, x, y, primary=True, button_size=new_button_size)
            
            button_font = ("Segoe UI", 14)  # Font family, size

            if os.name == 'nt':
                grade_button = tk.Button(self.content_frame, text="Grade", font=button_font, command=self.check_answers, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_button.place(x=1601, y=493)
            else:
                grade_button = tk.Button(self.content_frame, text="Grade",command=self.check_answers, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
                grade_button.place(x=823, y=260)

            finish_button = tk.Button(self.content_frame, font=button_font, text="Finish", command=self.get_scaled_scores, borderwidth=0, highlightthickness=0, bd=0, padx=0, pady=0)
            self.content_frame.update_idletasks()
            
            # Position the button in the bottom-right corner
            if os.name == 'nt':
                finish_button.place(x=1599, y=828)
            else:
                finish_button.place(x=15, y=260)
        
        def create_round_radiobutton(self, question, variable, value, x, y, primary=True, button_size=(15, 15)):
            # Paths for normal and selected state images
            png_normal_path = os.path.join(get_test_png_path(), f"{value.lower()}_locator_button.png")
            png_selected_path = os.path.join(get_test_png_path(), f"filled_locator_button.png")

            # Load and resize images
            normal_image = Image.open(png_normal_path).resize(button_size, Image.LANCZOS)
            selected_image = Image.open(png_selected_path).resize(button_size, Image.LANCZOS)
            
            normal_image = ImageTk.PhotoImage(normal_image)
            selected_image = ImageTk.PhotoImage(selected_image)
            
            def update_primary_buttons():
                for q in range(1, 16):
                    for ans in ['A', 'B', 'C', 'D', 'F', 'G', 'H', 'J']:
                        btn = self.primary_button_references[q].get(ans)
                        if btn and self.selected_answers[q].get() == ans:
                            btn.config(image=btn.selected_image)
                        elif btn:
                            btn.config(image=btn.normal_image)

            # Create a Radiobutton with the normal image
            radiobutton = tk.Radiobutton(self.content_frame, image=normal_image, variable=variable, value=value, 
                                        indicatoron=False, width=button_size[0], height=button_size[1], bd=0, 
                                        command=update_primary_buttons)
            radiobutton.normal_image = normal_image  # Keep references to the images
            radiobutton.selected_image = selected_image
            radiobutton.place(x=x, y=y)
            
            # Store a reference to the radiobutton in the appropriate dictionary
            if primary:
                self.primary_button_references[question][value] = radiobutton

        def enter_score(self):
            pass
        def get_scaled_scores(self):
            locator_raw = self.total_correct_answers_string_var.get()

            if locator_raw < 6:
                locator_result = 1
            elif 6 <= locator_raw < 9:
                locator_result = 2
            elif 9 <= locator_raw < 12:
                locator_result = 3
            else:
                locator_result = 4

            demo_message = (
            f"Student: {fields_dict['Student Name:'].get()}\n"
            f"Test Level: {fields_dict['Level:'].get()}\n"
            f"Test Form: {fields_dict['Form:'].get()}\n"
            f"Date: {fields_dict['Date:'].get()}\n")

            score_message = (
                f"Number Correct: {locator_raw}\n"
                f"Test to Administor: {locator_result}"
            )
            confirm_message = "\n\nDo you want to save the student's test data?"

            result_message = demo_message + "\n" + score_message + confirm_message
            if messagebox.askyesno("Save Test Results?", result_message):
                self.save_test()

        def save_test(self):
            # Get result from raw score:
            locator_raw = self.total_correct_answers_string_var.get()

            if locator_raw < 6:
                locator_result = 1
            elif 6 <= locator_raw < 9:
                locator_result = 2
            elif 9 <= locator_raw < 12:
                locator_result = 3
            else:
                locator_result = 4

            conn = sqlite3.connect(DB_FILE_PATH)
            cursor = conn.cursor()

            query = """
                INSERT INTO Testing (
                    student_id,
                    laces,
                    site,
                    staff,
                    primary_class,
                    date,
                    pre_or_post,
                    level,
                    form,
                    locator_answers,
                    locator_raw,
                    locator_result
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.execute(query, (
                fields_dict['Student ID'].get(),
                get_laces_id_for_save(),
                fields_dict['Site:'].get(),
                fields_dict['Aspire Staff:'].get(),
                fields_dict['Primary Class:'].get(),
                format_date(fields_dict['Date:'].get().replace('/', '-'), out=True),
                fields_dict['Pre/Post:'].get(),
                fields_dict['Level:'].get(),
                fields_dict['Form:'].get(),
                "".join([v.get() if v.get() != "" else " " for v in self.selected_answers.values()]),
                self.total_correct_answers_string_var.get(),
                locator_result
            ))
            conn.commit()
            conn.close()

            # Repopulate TestingTab treeview
            if on_save:
                on_save()

        def check_answers(self):
            correct_answers = list('CGAFCJCFCGCHCFB')
            # Dictionary of question number, student's response, and correct answer
            student_responses = {}

            # Grade locator questions
            for i, (question, answer_choice) in enumerate(self.selected_answers.items()):
                if i == 0:
                    self.total_correct_answers_string_var.set(0)
                answer_choice = answer_choice.get()
                correct_answer = correct_answers[i]
                student_responses[question] = (answer_choice, correct_answer)
                if answer_choice == correct_answer:
                    old_score = self.total_correct_answers_string_var.get()
                    new_score = old_score + 1
                    self.total_correct_answers_string_var.set(new_score)
    
            # Destroy existing labels before creating new ones
            if hasattr(self, 'grade_reading_label') and self.grade_reading_label.winfo_exists():
                self.grade_reading_label.destroy()

            # Display Total label if score exists
            if self.total_correct_answers_string_var.get():
                self.grade_reading_label = tk.Label(self.content_frame, textvariable=self.total_correct_answers_string_var)
                if os.name == 'nt':
                    x_position = 1627 if int(self.total_correct_answers_string_var.get()) < 10 else (1627 - 10)
                    self.grade_reading_label.place(x=x_position, y=416)
                    self.grade_reading_label.config(font=("Segoe UI", 23))
                else:
                    x_position = 848 if int(self.total_correct_answers_string_var.get()) < 10 else 842
                    self.grade_reading_label.place(x=x_position, y=218)
                    self.grade_reading_label.config(font=("Segoe UI", 24))

            # Calculate scaled score
            
        def grade_test(self):
            self.answer_sheet_window = tk.Toplevel(self.add_test_window)
            self.answer_sheet_window.title((
                f'Level {fields_dict["Level:"].get()}, Form {fields_dict["Form:"].get()} - {fields_dict["Student Name:"].get()} - {fields_dict["Date:"].get()}')
            )
            
            # Setup canvas and scrollbars
            canvas = tk.Canvas(self.answer_sheet_window, width=1511, height=800)
            v_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.VERTICAL, command=canvas.yview)
            h_scrollbar = tk.Scrollbar(self.answer_sheet_window, orient=tk.HORIZONTAL, command=canvas.xview)
            
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            # h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create a frame inside the canvas to hold all content
            self.content_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
            
            # Ensure the scroll region matches the content size
            self.content_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))

            # Bind the right-click event (Button-3 in Tkinter) to the function
            self.answer_sheet_window.bind("<Button-3>", right_click)
            canvas.bind_all("<MouseWheel>", lambda e: on_mouse_wheel(e, canvas))

            # Label to display the test image
            self.image_label = tk.Label(self.content_frame)
            self.image_label.pack()

            center_window(self.answer_sheet_window, 1800, 900)
            # self.answer_sheet_window.state("zoomed")
            # self.answer_sheet_window.resizable(False, False)

            # Update the display with the constant image and size
            self.update_test_display()
            

    def check_fields():
        selected_index = combobox.current()
        if all(var.get().strip() for key, var in fields_dict.items() if key not in ('LACES #:')):
        # if all(var.get().strip() for key, var in fields_dict.items()):
            grade_button.config(state='normal')
            enter_score_button.config(state='normal')
        else:
            grade_button.config(state='disabled')
            enter_score_button.config(state='disabled')
        if fields_dict["Level:"].get() == "Locator":
            fields_dict["Form:"].set("Locator")
        if fields_dict["Form:"].get() == "Locator" and fields_dict["Level:"].get() != "Locator":
            fields_dict["Form:"].set("")

    def load_test():
        if fields_dict["Level:"].get() in ['2', '3', '4']:
            LevelTwoTest(add_test_window).grade_test()
        elif fields_dict["Level:"].get() == '1':
            LevelOneTest(add_test_window).grade_test()
        elif fields_dict["Level:"].get() == 'Locator':
            LocatorTest(add_test_window).grade_test()

    def save_manual_score(reading_scaled, reading_nrs, writing_scaled, writing_nrs, combined, current_test_date):
        # First, determine if student got MSG
        conn = sqlite3.connect(DB_FILE_PATH)
        cursor = conn.cursor()
        query = """
            SELECT
                reading_nrs,
                writing_nrs,
                MAX(date)
            FROM
                Testing
            WHERE
                student_id = ?
                AND date < ?
        """
        cursor.execute(query, (fields_dict['Student ID'].get(), format_date(current_test_date, out=True)))
        last_test = cursor.fetchone()
        conn.close()

        if last_test[0]:
            # Handle possible None values in the current reading/writing scores
            if reading_nrs is not None:
                current_reading_nrs = int(str(reading_nrs).replace('+', ''))
            else:
                current_reading_nrs = None  # If reading_nrs is None, keep it as None
            
            if writing_nrs is not None:
                current_writing_nrs = int(str(writing_nrs).replace('+', ''))
            else:
                current_writing_nrs = None  # If writing_nrs is None, keep it as None

            # Handle possible None values in the last test's reading/writing scores
            if last_test[0] is not None:
                if "+" in str(last_test[0]):
                    last_reading_nrs = int(str(last_test[0]).replace('+', '')) + 1
                else:
                    last_reading_nrs = int(str(last_test[0]))
            else:
                last_reading_nrs = None  # If last_test[0] is None, keep it as None
            
            if last_test[1] is not None:
                if '+' in str(last_test[1]):
                    last_writing_nrs = int(str(last_test[1]).replace('+', '')) + 1
                else:
                    last_writing_nrs = int(str(last_test[1]))
            else:
                last_writing_nrs = None  # If last_test[1] is None, keep it as None

            # Initialize message area and got_msg
            msg_area = ""
            got_msg = 0

            # Compare reading scores only if both current and last reading scores are not None
            if current_reading_nrs is not None and last_reading_nrs is not None:
                if current_reading_nrs > last_reading_nrs:
                    got_msg = 1
                    msg_area += f'Reading: {last_reading_nrs} -> {current_reading_nrs}'
            
            # Compare writing scores only if both current and last writing scores are not None
            if current_writing_nrs is not None and last_writing_nrs is not None:
                if current_writing_nrs > last_writing_nrs:
                    got_msg = 1
                    if msg_area != "":
                        msg_area += f'\nWriting: {last_writing_nrs} -> {current_writing_nrs}'    
                    else:
                        msg_area += f'Writing: {last_writing_nrs} -> {current_writing_nrs}'
        else:
            got_msg = 0
            msg_area = ""

        # Get correct date string
        # formatted_date = format_date(fields_dict['Date:'].get().replace('/', '-'), out=True)
        # print(formatted_date)
        conn = sqlite3.connect(DB_FILE_PATH)
        cursor = conn.cursor()

        query = """
            INSERT INTO Testing (
                student_id,
                laces,
                site,
                staff,
                primary_class,
                date,
                pre_or_post,
                level,
                form,
                reading_scaled,
                reading_nrs,
                writing_scaled,
                writing_nrs,
                combined,
                got_msg,
                msg_area
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        cursor.execute(query, (
            fields_dict['Student ID'].get(),
            get_laces_id_for_save(),
            fields_dict['Site:'].get(),
            fields_dict['Aspire Staff:'].get(),
            fields_dict['Primary Class:'].get(),
            format_date(fields_dict['Date:'].get().replace('/', '-'), out=True),
            fields_dict['Pre/Post:'].get(),
            fields_dict['Level:'].get(),
            fields_dict['Form:'].get(),
            reading_scaled,
            reading_nrs,
            writing_scaled,
            writing_nrs,
            combined,
            got_msg,
            msg_area
        ))
        conn.commit()
        conn.close()

        # Repopulate TestingTab treeview
        if on_save:
            on_save()
            add_test_window.destroy()

    # Function to display a pop-up window with options to re-enter the score or choose 'no score'
    def handle_missing_score(score_type, key, score):
        # result_score = None
        nrs_value = None


        return score, nrs_value
        pass

    # Function to get the score and handle missing cases
    def get_score_or_prompt(score_type, key, score):
        try:
            # Try to find the score index in the data
            score_index = tabe_scoring_tables[key]['scale_score'].index(score)
            return score, tabe_scoring_tables[key]['NRS'][score_index]
        except ValueError:
            # If score not found, call the popup handler
            return handle_missing_score(score_type, key, score)

    # Existing function integrated with the new logic
    def grade_manual_score(reading, writing, enter_manual_score_window):
        # Access the correct section of the data using Level and Form
        reading_level_form_key = f"Reading {fields_dict['Level:'].get()}{fields_dict['Form:'].get()}"
        writing_level_form_key = f"Writing {fields_dict['Level:'].get()}{fields_dict['Form:'].get()}"

        # Get reading and writing scores, or prompt if not found
        reading_score, reading_nrs = get_score_or_prompt("Reading", reading_level_form_key, reading)
        writing_score, writing_nrs = get_score_or_prompt("Writing", writing_level_form_key, writing)

        # Calculate the combined score
        if reading_score is None or writing_score is None:
            combined = None
        else:
            combined = int(((reading_score + writing_score) / 2))

        # Prepare demo message
        demo_message = (
            f"Student: {fields_dict['Student Name:'].get()}\n"
            f"Test Level: {fields_dict['Level:'].get()}\n"
            f"Test Form: {fields_dict['Form:'].get()}\n"
            f"Date: {fields_dict['Date:'].get()}\n"
        )

        # Prepare the score message
        scaled_scores = {
            "Reading Scaled": reading_score if reading_score is not None else "NULL",
            "Reading NRS": reading_nrs if reading_nrs is not None else "NULL",
            "Writing Scaled": writing_score if writing_score is not None else "NULL",
            "Writing NRS": writing_nrs if writing_nrs is not None else "NULL",
            "Combined": combined if combined is not None else "NULL"
        }

        # Display the results and ask for confirmation to save
        score_message = "\n".join([f"{k}: {v}" for k, v in scaled_scores.items()])
        confirm_message = "\n\nDo you want to save the student's test data?"

        result_message = demo_message + score_message + confirm_message
        if messagebox.askyesno("Save Test Results?", result_message):
            save_manual_score(reading_score, reading_nrs, writing_score, writing_nrs, combined, fields_dict['Date:'].get())
            enter_manual_score_window.destroy()
    
    def enter_manual_score():
        enter_manual_score_window = tk.Toplevel()
        enter_manual_score_window.title("Enter Manual Score")

        enter_reading_scale_score = IntVar()
        enter_writing_scale_score = IntVar()

        reading_score_label = tk.Label(enter_manual_score_window, text="Reading Scale Score:", font=default_font)
        reading_score_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        reading_score_entry = ttk.Entry(enter_manual_score_window, textvariable=enter_reading_scale_score, width=15, font=default_font)
        reading_score_entry.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

        writing_score_label = tk.Label(enter_manual_score_window, text="Writing Scale Score:", font=default_font)
        writing_score_label.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
        
        writing_score_entry = ttk.Entry(enter_manual_score_window, textvariable=enter_writing_scale_score, width=15, font=default_font)
        writing_score_entry.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')

        grade_button = Button(enter_manual_score_window, text="Grade", command=lambda: grade_manual_score(enter_reading_scale_score.get(), enter_writing_scale_score.get(), enter_manual_score_window), font=default_font)
        grade_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

        enter_manual_score_window.grid_rowconfigure(2, weight=1)
        enter_manual_score_window.grid_columnconfigure(1, weight=1)

    def get_student_laces_id(student_id):
        """Return Student.laces_id for the selected student, or an empty string."""
        if not student_id:
            return ""

        conn = sqlite3.connect(DB_FILE_PATH)
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT laces_id
            FROM Student
            WHERE student_id = ?
            """,
            (student_id,)
        )

        row = cursor.fetchone()
        conn.close()

        if row and row[0] is not None:
            return str(row[0]).strip()

        return ""


    def save_laces_id_to_student_if_missing(student_id, laces_id):
        """
        Save the typed LACES ID to Student.laces_id only if Student.laces_id
        is currently NULL or blank.
        """
        student_id = str(student_id).strip()
        laces_id = str(laces_id).strip()

        if not student_id or not laces_id:
            return

        conn = sqlite3.connect(DB_FILE_PATH)
        cursor = conn.cursor()

        cursor.execute(
            """
            UPDATE Student
            SET laces_id = ?
            WHERE student_id = ?
              AND (laces_id IS NULL OR TRIM(laces_id) = '')
            """,
            (laces_id, student_id)
        )

        conn.commit()
        conn.close()


    def get_laces_id_for_save():
        """
        Use Student.laces_id as the source of truth.
        If Student.laces_id is blank and the user typed a LACES ID,
        save the typed value back to Student.laces_id and return it.
        """
        student_id = fields_dict['Student ID'].get()
        typed_laces_id = fields_dict['LACES #:'].get().strip()

        existing_laces_id = get_student_laces_id(student_id)

        if existing_laces_id:
            fields_dict['LACES #:'].set(existing_laces_id)
            return existing_laces_id

        if typed_laces_id:
            save_laces_id_to_student_if_missing(student_id, typed_laces_id)
            return typed_laces_id

        return ""


    def set_student_and_laces_id(combo_ind):
        student_id = student_data[combo_ind][0]
        fields_dict["Student ID"].set(student_id)

        # Prefer the laces_id that came from get_student_data().
        # student_data row layout:
        #   0 = student_id
        #   1 = student name
        #   2 = site
        #   3 = laces_id
        if len(student_data[combo_ind]) > 3 and student_data[combo_ind][3] is not None:
            laces_id = str(student_data[combo_ind][3]).strip()
        else:
            laces_id = get_student_laces_id(student_id)

        fields_dict['LACES #:'].set(laces_id)

        fields_dict['LACES #:'].trace('w', lambda name, index, mode, var=string_var: check_fields())


    # Main window setup
    add_test_window = tk.Toplevel()
    add_test_window.title("Add Test")
    add_test_window.withdraw()
    center_window(add_test_window, 750, 900)

    # Populate student data
    student_data = get_student_data(DB_FILE_PATH)
    sites = sorted(list(set([s[2] if s[2] else '' for s in student_data])))
    students = [s[1] for s in student_data]
    fields = [
        ("Pre/Post:", StringVar()),
        ("Site:", StringVar()),
        ("Aspire Staff:", StringVar()),
        ("Primary Class:", StringVar(value="ESOL")),
        ("Student Name:", StringVar()),
        ("Student ID", StringVar()),
        ("LACES #:", StringVar()),
        ("Level:", StringVar()),
        ("Form:", StringVar()),
        ("Date:", StringVar())
    ]
    fields_dict = {key: value for key, value in fields}

    for i, (label_text, string_var) in enumerate(fields):
        if label_text in ['Pre/Post:', 'Site:', 'Primary Class:', 'Student Name:', 'Level:', 'Form:']:
            label = tk.Label(add_test_window, text=label_text, font=default_font)
            label.grid(row=i, column=0, padx=10, pady=10, sticky='nsew')

            if label_text == "Pre/Post:":
                combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=["Pre", "Post"], width=15, font=default_font)
                combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
            elif label_text == "Site:":
                combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=sites, width=15, font=default_font)
                combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
            elif label_text == "Primary Class:":
                combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=['ESOL','CIT', 'ESOL/CIT', 'IET'], width=15, font=default_font)
                combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
            elif label_text == "Student Name:":
                name_combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=students, width=15, font=default_font)
                name_combobox.bind("<<ComboboxSelected>>", lambda event: set_student_and_laces_id(name_combobox.current()))
                name_combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')     
            elif label_text == "Level:":
                combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=["Locator", 1, 2, 3, 4, 'TABE 12/13'], width=15, font=default_font)
                combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
            elif label_text == "Form:":
                combobox = AutocompleteCombobox(add_test_window, textvariable=string_var, values=['C', 'D'], width=15, font=default_font)
                combobox.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
        elif label_text in ['Aspire Staff:', 'LACES #:']:
            label = tk.Label(add_test_window, text=label_text, font=default_font)
            label.grid(row=i, column=0, padx=10, pady=10, sticky='nsew')
            entry = ttk.Entry(add_test_window, textvariable=string_var, width=15, font=default_font)
            entry.grid(row=i, column=1, padx=10, pady=10, sticky='nsew')
        elif label_text == "Date:":
            label = tk.Label(add_test_window, text=label_text, font=default_font)
            label.grid(row=i, column=0, padx=10, pady=10, sticky='nsew')
            large_font = font.Font(family="Helvetica", size=14, weight="bold")
            cal = Calendar(add_test_window, selectmode='day', 
                                year=datetime.today().year, month=datetime.today().month, 
                                day=datetime.today().day, font=large_font, 
                                headersbackground="lightblue",
                                background="lightgrey", foreground="black",
                                selectbackground="blue", selectforeground="white")
            cal.grid(row=i, column=1, columnspan=4, padx=10, pady=5)
            
            def update_date(event):
                string_var.set(cal.get_date())
            
            cal.bind("<<CalendarSelected>>", update_date)
        string_var.trace('w', lambda name, index, mode, var=string_var: check_fields())

    grade_button = tk.Button(add_test_window, text="Grade Test", command=load_test, font=default_font)
    grade_button.grid(row=len(fields) + 1, column=1, padx=10, pady=10)
    grade_button.config(state='disabled')

    # Create manual entry button
    enter_score_button = tk.Button(add_test_window, text="Enter Score",command=enter_manual_score, font=default_font)
    enter_score_button.grid(row=len(fields)+1, column=2, padx=10, pady=5)
    grade_button.config(state='disabled')

    # Set initial state of "Form" field to blank
    fields_dict["Form:"].set("")
    add_test_window.deiconify()

# if __name__ == "__main__":
#     run_test_grading_gui()
