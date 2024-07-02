import customtkinter as ctk
import matplotlib.pyplot as plt
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

class DataEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Entry Form")
        self.root.geometry("1200x800")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.students_data = []
        self.current_student = None
        self.student_counter = 0

        # Create main frame
        self.main_frame = ctk.CTkFrame(root, corner_radius=15)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Define the fields
        self.fields = ["Name", "Meal Time", "School Hours", "Tuition", "Outdoor Games",
                       "Homework/ Self Study time", "Family Time", "Online Games",
                       "On phonecall", "TV", "Sleep hours", "Any other", "Social Media"]

        self.entries = {}

        # Create entry widgets for each field
        for idx, field in enumerate(self.fields):
            label = ctk.CTkLabel(self.main_frame, text=field)
            label.grid(row=idx, column=0, padx=10, pady=5, sticky="w")
            entry = ctk.CTkEntry(self.main_frame, width=200, corner_radius=15)
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="w")
            self.entries[field] = entry

        # Add a listbox for days with the theme
        self.days_listbox = ctk.CTkTextbox(self.main_frame, height=10, width=300, corner_radius=15)
        self.days_listbox.grid(row=0, column=2, rowspan=len(self.fields), padx=10, pady=5, sticky="ns")

        self.students_textbox = ctk.CTkTextbox(self.main_frame, height=20, width=300, corner_radius=15)
        self.students_textbox.grid(row=0, column=3, rowspan=len(self.fields)+2, padx=10, pady=5, sticky="ns")

        # Create buttons
        self.add_day_button = ctk.CTkButton(self.main_frame, text="Add Tracker", command=self.add_day, corner_radius=15)
        self.add_day_button.grid(row=len(self.fields), column=0, pady=10, sticky="w")

        self.new_student_button = ctk.CTkButton(self.main_frame, text="Add New Student", command=self.add_new_student, corner_radius=15)
        self.new_student_button.grid(row=len(self.fields), column=1, pady=10, sticky="e")

        self.pie_button = ctk.CTkButton(self.main_frame, text="Visualize Pie Chart", command=self.visualize_pie_chart, corner_radius=15)
        self.pie_button.grid(row=len(self.fields)+1, column=0, pady=20, sticky="w")

        self.bar_button = ctk.CTkButton(self.main_frame, text="Visualize Bar Graph", command=self.visualize_bar_graph, corner_radius=15)
        self.bar_button.grid(row=len(self.fields)+1, column=1, pady=20, sticky="e")

        self.save_button = ctk.CTkButton(self.main_frame, text="Save to Excel", command=self.save_to_excel, corner_radius=15)
        self.save_button.grid(row=len(self.fields)+1, column=2, pady=20, sticky="n")

        self.exit_button = ctk.CTkButton(self.main_frame, text="Exit", command=self.exit_app, corner_radius=15)
        self.exit_button.grid(row=len(self.fields)+1, column=4, pady=20, sticky="e", padx=10)

        self.days_data = []
        self.name_set = False

        # Add student count label
        self.student_count_label = ctk.CTkLabel(self.main_frame, text="Total Students: 0", font=("Arial", 16))
        self.student_count_label.grid(row=len(self.fields)+2, column=0, columnspan=5, pady=10)

    def add_day(self):
        day_data = {}
        student_name = self.entries["Name"].get().strip()
        if not student_name:
            messagebox.showwarning("Missing Name", "Please enter the student's name")
            return

        if self.current_student and self.current_student != student_name:
            self.refresh()

        self.current_student = student_name
        for field in self.fields:
            value = self.entries[field].get()
            if field == "Name":
                day_data[field] = student_name
            else:
                try:
                    day_data[field] = float(value)
                except ValueError:
                    messagebox.showwarning("Invalid input", f"Please enter a valid number for {field}")
                    return

        day_data["tracker_number"] = len(self.days_data) + 1
        self.students_data.append(day_data)
        self.days_data.append(day_data)
        self.days_listbox.insert(tk.END, f"Tracker {len(self.days_data)} added for {student_name}\n")
        self.students_textbox.insert(tk.END, f"{student_name}: {day_data}\n")
        self.clear_entries()
        self.update_student_count()

    def clear_entries(self):
        for field, entry in self.entries.items():
            if field != "Name":
                entry.delete(0, tk.END)

    def get_avg_data(self, data):
        if not data:
            return None

        avg_data = {field: 0 for field in self.fields if field != "Name"}
        for day_data in data:
            for field in avg_data:
                avg_data[field] += day_data[field]
        avg_data = {field: value / len(data) for field, value in avg_data.items()}
        return avg_data

    def save_graph(self, fig):
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
        if file_path:
            fig.savefig(file_path)
            messagebox.showinfo("Success", f"Graph saved successfully at {file_path}")

    def visualize_pie_chart(self):
        avg_data = self.get_avg_data(self.days_data)
        if avg_data:
            name = self.current_student
            values = list(avg_data.values())
            labels = list(avg_data.keys())
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
            ax.set_title(f'Activity Distribution for {name}')
            plt.show()
            self.save_graph(fig)

    def visualize_bar_graph(self):
        avg_data = self.get_avg_data(self.days_data)
        if avg_data:
            name = self.current_student
            values = list(avg_data.values())
            labels = list(avg_data.keys())
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.bar(labels, values, color=plt.cm.Paired.colors)
            ax.set_xticklabels(labels, rotation=45, ha='right')
            ax.set_title(f'Activity Distribution for {name}')
            ax.set_xlabel('Activities')
            ax.set_ylabel('Time Spent')
            plt.tight_layout()
            plt.show()
            self.save_graph(fig)

    def refresh(self):
        self.days_data = []
        self.days_listbox.delete("1.0", tk.END)
        self.students_textbox.delete("1.0", tk.END)
        self.clear_entries()
        self.current_student = None

    def add_new_student(self):
        self.refresh()
        self.student_counter += 1
        self.update_student_count()

    def exit_app(self):
        self.root.destroy()

    def update_student_count(self):
        total_students = len(set(student["Name"] for student in self.students_data))
        self.student_count_label.configure(text=f"Total Students: {total_students}")

    def save_to_excel(self):
        if not self.students_data:
            messagebox.showwarning("No data", "No data to save")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            all_data = []
            current_name = None
            tracker_number = 1
            for idx, day_data in enumerate(self.students_data):
                if current_name != day_data["Name"]:
                    if current_name is not None:
                        all_data.append({field: "" for field in self.fields + ["tracker_number"]})  # Add a blank row between different students
                    current_name = day_data["Name"]
                    tracker_number = 1  # Reset tracker number for a new student

                day_data["tracker_number"] = tracker_number
                day_data["student_number"] = idx + 1
                all_data.append(day_data)
                tracker_number += 1

            df = pd.DataFrame(all_data)
            df = df[["student_number", "Name", "tracker_number"] + self.fields[1:]]
            df.to_excel(file_path, sheet_name='Student Data', index=False)

            # Adding bar graph for average data of all students
            avg_data = self.get_avg_data(self.students_data)
            if avg_data:
                values = list(avg_data.values())
                labels = list(avg_data.keys())
                fig, ax = plt.subplots(figsize=(12, 6))
                ax.bar(labels, values, color=plt.cm.Paired.colors)
                ax.set_xticklabels(labels, rotation=45, ha='right')
                ax.set_title('Average Activity Distribution of All Students')
                ax.set_xlabel('Activities')
                ax.set_ylabel('Time Spent')
                plt.tight_layout()

                # Save the figure to a BytesIO object
                from io import BytesIO
                img_data = BytesIO()
                fig.savefig(img_data, format='png')
                img_data.seek(0)

                # Insert image into Excel
                from openpyxl import load_workbook
                from openpyxl.drawing.image import Image
                workbook = load_workbook(file_path)
                worksheet = workbook['Student Data']
                img = Image(img_data)
                worksheet.add_image(img, 'N1')
                workbook.save(file_path)

            messagebox.showinfo("Success", f"Data saved successfully at {file_path}")

# Main function to run the application
if __name__ == "__main__":
    root = ctk.CTk()
    app = DataEntryApp(root)
    root.mainloop()
