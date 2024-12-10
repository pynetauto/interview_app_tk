"""
Author: Brendan (Byong Chol) Choi
App Name: italchemy_interview_app_v1.0.py
Date: 09/Dec/2024
Purpose: 
This application helps interviewers prepare questions and answers, ensuring 
consistent questions are asked for all candidates applying for a role. 

Instructions:
1. Upload a set of interview questions in 'interview_questions_v1.xlsx'.
2. Run the compiled program along with the .xlsx file in the same folder 
   during the interview.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import random

# For compiled executable or running script
if getattr(sys, 'frozen', False):
    excel_file = os.path.join(sys._MEIPASS, 'interview_questions_v1.0.xlsx')
    logo_file = os.path.join(sys._MEIPASS, 'italchemy_logo.png')
else:
    excel_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'interview_questions_v1.0.xlsx')
    logo_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'italchemy_logo.png')

# Check if the Excel file exists
if not os.path.exists(excel_file):
    messagebox.showerror("Error", "Excel file not found.")
    sys.exit()

# Check if the logo file exists
if not os.path.exists(logo_file):
    messagebox.showerror("Error", "Logo file not found.")
    sys.exit()

# Try to load the Excel file and handle errors if the file is missing or invalid
try:
    df = pd.read_excel(excel_file, engine='openpyxl')  # Using openpyxl to read .xlsx files
except Exception as e:
    messagebox.showerror("Error", f"Failed to load Excel file: {e}")
    sys.exit()

# Validate that the required columns are present
required_columns = ['Question Number', 'Interview Question', 'Topic', 'Answer', 'Difficulty']
if not all(col in df.columns for col in required_columns):
    messagebox.showerror("Error", "Missing one or more required columns in the Excel file.")
    sys.exit()

# Clean the Difficulty column (strip spaces and convert to lowercase)
df['Difficulty'] = df['Difficulty'].str.strip().str.lower()

# Pre-process and store questions based on difficulty levels to avoid repeated filtering
easy_questions = df[df['Difficulty'] == 'easy'].reset_index(drop=True)
medium_questions = df[df['Difficulty'] == 'medium'].reset_index(drop=True)
hard_questions = df[df['Difficulty'] == 'hard'].reset_index(drop=True)

# Track which questions have been used for each difficulty
used_easy = set()
used_medium = set()
used_hard = set()

# Ensure we have enough questions to sample
def get_random_questions(df, n=3, used_questions=set()):
    remaining_questions = df[~df['Question Number'].isin(used_questions)]
    if len(remaining_questions) < n:
        return remaining_questions, used_questions  # Return all remaining questions if fewer than 'n'
    selected_questions = remaining_questions.sample(n=n, random_state=random.randint(1, 1000))
    used_questions.update(selected_questions['Question Number'])  # Mark these as used
    return selected_questions, used_questions

# Select random questions from each difficulty level
def refresh_questions(difficulty):
    global random_easy, random_medium, random_hard
    global used_easy, used_medium, used_hard

    if difficulty == 'easy':
        random_easy, used_easy = get_random_questions(easy_questions, 3, used_easy)
    elif difficulty == 'medium':
        random_medium, used_medium = get_random_questions(medium_questions, 3, used_medium)
    elif difficulty == 'hard':
        random_hard, used_hard = get_random_questions(hard_questions, 3, used_hard)

# Function to format questions, topics, and answers for display
def format_questions(questions):
    formatted = ""
    for idx, row in questions.iterrows():
        formatted += f"Q{row['Question Number']}: {row['Interview Question']}\n"
        formatted += f"Topic: {row['Topic']}\n"  # Display Topic as the keywords
        formatted += f"A: {row['Answer']}\n\n"  # Changed "Answer: " to "A: "
    return formatted

# Function to display questions based on difficulty
def show_questions(difficulty, button):
    global random_easy, random_medium, random_hard
    global used_easy, used_medium, used_hard

    refresh_questions(difficulty)  # Refresh the random questions based on selected difficulty
    
    if difficulty == 'easy':
        questions_text = format_questions(random_easy)
        difficulty_label.config(text="Difficulty: Easy", fg="green")  # Update the difficulty label
    elif difficulty == 'medium':
        questions_text = format_questions(random_medium)
        difficulty_label.config(text="Difficulty: Medium", fg="blue")  # Update the difficulty label
    elif difficulty == 'hard':
        questions_text = format_questions(random_hard)
        difficulty_label.config(text="Difficulty: Hard", fg="maroon")  # Update the difficulty label
    else:
        print("You are my lucky star!")

    # Efficiently update the Text widget with the selected questions
    result_text.replace(1.0, tk.END, questions_text)  # Using replace for efficient update

    # Highlight the pressed button and reset others
    highlight_button(button)

    # Update the question counts after displaying the questions
    update_question_counts()

# Function to change the button text color to red and reset others
def highlight_button(pressed_button):
    # Reset all buttons' text color to green
    easy_button.config(fg="green")
    medium_button.config(fg="blue")
    hard_button.config(fg="maroon")
    
    # Change the font color of the pressed button to red
    pressed_button.config(fg="red")

# Reset function to re-enable the buttons and reset all used questions for each level
def reset_level(level):
    global used_easy, used_medium, used_hard
    if level == 'easy':
        used_easy.clear()
        easy_button.config(fg="green")  # Reset text color to green
    elif level == 'medium':
        used_medium.clear()
        medium_button.config(fg="blue")  # Reset text color to blue
    elif level == 'hard':
        used_hard.clear()
        hard_button.config(fg="maroon")  # Reset text color to maroon

    result_text.delete(1.0, tk.END)  # Clear the text area
    update_question_counts()  # Update question counts after reset

# Zoom functionality to adjust font size
def zoom_in():
    current_font = result_text.cget("font")
    size = int(current_font.split()[1]) + 2  # Increase font size
    result_text.config(font=("Arial", size))

def zoom_out():
    current_font = result_text.cget("font")
    size = max(8, int(current_font.split()[1]) - 2)  # Decrease font size, minimum 8
    result_text.config(font=("Arial", size))

# Function to update the question count display below the buttons
def update_question_counts():
    total_easy = len(easy_questions)
    total_medium = len(medium_questions)
    total_hard = len(hard_questions)

    remaining_easy = total_easy - len(used_easy)
    remaining_medium = total_medium - len(used_medium)
    remaining_hard = total_hard - len(used_hard)

    # Update labels with the question count (remaining / total)
    easy_count_label.config(text=f"Easy Qs: {remaining_easy}/{total_easy}")
    medium_count_label.config(text=f"Medium Qs: {remaining_medium}/{total_medium}")
    hard_count_label.config(text=f"Hard Qs: {remaining_hard}/{total_hard}")

# Create the main window with increased size
root = tk.Tk()
root.title("Random Interview Questions")

# Set the window size (width x height)
root.geometry("1300x800")  # Increased width to make the display area 1.5 times wider

# Add a label to display the selected questions with larger font size
result_var = tk.StringVar()

# Create a frame to hold the Text widget and scrollbar
frame = tk.Frame(root)
frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

# Create the Text widget with a scrollbar for displaying questions
result_text = tk.Text(frame, wrap=tk.WORD, font=("Arial", 16), height=20)
result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Add a scrollbar to the Text widget
scrollbar = tk.Scrollbar(frame, command=result_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
result_text.config(yscrollcommand=scrollbar.set)

# Add a frame for buttons at the top center
button_frame = tk.Frame(root)
button_frame.pack(pady=10, padx=20)

# Add buttons for each difficulty level and reset button
easy_button = tk.Button(button_frame, text="Easy", command=lambda: show_questions('easy', easy_button), font=("Arial", 12), width=10)
medium_button = tk.Button(button_frame, text="Medium", command=lambda: show_questions('medium', medium_button), font=("Arial", 12), width=10)
hard_button = tk.Button(button_frame, text="Hard", command=lambda: show_questions('hard', hard_button), font=("Arial", 12), width=10)

# Add Reset buttons for each difficulty level under their respective buttons
reset_easy_button = tk.Button(button_frame, text="Reset-E", command=lambda: reset_level('easy'), font=("Arial", 12), width=10)
reset_medium_button = tk.Button(button_frame, text="Reset-M", command=lambda: reset_level('medium'), font=("Arial", 12), width=10)
reset_hard_button = tk.Button(button_frame, text="Reset-H", command=lambda: reset_level('hard'), font=("Arial", 12), width=10)

easy_button.grid(row=0, column=0, padx=5)
medium_button.grid(row=0, column=1, padx=5)
hard_button.grid(row=0, column=2, padx=5)
reset_easy_button.grid(row=1, column=0, padx=5)
reset_medium_button.grid(row=1, column=1, padx=5)
reset_hard_button.grid(row=1, column=2, padx=5)

# Add Zoom buttons to the right edge
zoom_in_button = tk.Button(button_frame, text="Zoom In", command=zoom_in, font=("Arial", 12), width=10)
zoom_out_button = tk.Button(button_frame, text="Zoom Out", command=zoom_out, font=("Arial", 12), width=10)

zoom_in_button.grid(row=2, column=1, padx=5)
zoom_out_button.grid(row=2, column=2, padx=5)

# Add Difficulty Level Indicator label
difficulty_label = tk.Label(root, text="Difficulty: Not Selected", font=("Arial", 12), fg="gray")
difficulty_label.pack(pady=10)

# Add a frame to hold the question count labels below the buttons
count_frame = tk.Frame(root)
count_frame.pack(pady=10)

# Add labels for question counts below the buttons
easy_count_label = tk.Label(count_frame, text="Easy Qs: 0/0", font=("Arial", 12), fg="green")
easy_count_label.grid(row=0, column=0, pady=5)

medium_count_label = tk.Label(count_frame, text="Medium Qs: 0/0", font=("Arial", 12), fg="blue")
medium_count_label.grid(row=0, column=1, pady=5)

hard_count_label = tk.Label(count_frame, text="Hard Qs: 0/0", font=("Arial", 12), fg="maroon")
hard_count_label.grid(row=0, column=2, pady=5)

# Create the logo label and position it at the bottom-left corner
logo = tk.PhotoImage(file=logo_file)  # Load the logo image
logo_label = tk.Label(root, image=logo)
logo_label.place(x=20, y=700, anchor="w")  # Position it at the bottom-left corner

# Initial setup
update_question_counts()

# Run the Tkinter main loop
root.mainloop()
