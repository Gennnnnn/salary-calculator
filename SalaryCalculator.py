import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import json
import pandas as pd
from datetime import datetime  # Import datetime module

class SalaryCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Salary Calculator")
        self.root.state("zoomed")
        root.resizable(True, True)
        self.data = []
        
        # Set the desired folder path
        self.base_folder = r"C:\SCfiles\Workers"
        self.receipts_folder = r"C:\SCfiles\receipts"

        os.makedirs(self.base_folder, exist_ok=True)
        os.makedirs(self.receipts_folder, exist_ok=True)

        self.json_file_path = os.path.join(self.base_folder, "salary_data.json")
          
        self.finishing_employees = []
        self.carpentry_employees = []
        
        # Load existing data
        self.load_data()
        self.create_widgets()
    
    def center_window(self, width, height, window=None):
        window = window or self.root
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        # Set background color for the main window
        self.root.configure(bg="#f5f5dc")  # Light beige background

        # Title Label
        tk.Label(self.root, text="Salary Calculator", font=("Arial", 30, "bold"), fg="green", bg="#f5f5dc").pack(pady=20)

        frame_top = tk.Frame(self.root, bg="#f5f5dc")
        frame_top.pack(pady=10)

        # Finishing Section
        frame_finishing = tk.LabelFrame(frame_top, text="Finishing", fg="green", font=("Arial", 24), bg="#f5f5dc")
        frame_finishing.pack(side="left", padx=20, pady=10, fill="both", expand=True)

        ttk.Label(frame_finishing, text="Select Worker:", font=("Arial", 20), background="#f5f5dc").pack(padx=10, pady=10)

        self.finishing_employee_dropdown = ttk.Combobox(frame_finishing, state="readonly", font=("Arial", 18), width=40)
        self.finishing_employee_dropdown.pack(padx=10)
        self.finishing_employee_dropdown['values'] = ["------"] + [emp[0] for emp in self.finishing_employees]
        self.finishing_employee_dropdown.current(0)
        self.finishing_employee_dropdown.bind("<<ComboboxSelected>>", lambda event: self.handle_employee_selection("Finishing"))

        # Worker Button (Finishing)
        tk.Button(frame_finishing, text="Worker", command=lambda: self.open_employee_management("Finishing"), 
                bg="#4CAF50", fg="white", font=("Arial", 18, "bold"), width=20).pack(pady=10)  # Green Button

        # Carpentry Section
        frame_carpentry = tk.LabelFrame(frame_top, text="Carpentry", fg="green", font=("Arial", 24), bg="#f5f5dc")
        frame_carpentry.pack(side="right", padx=20, pady=10, fill="both", expand=True)

        ttk.Label(frame_carpentry, text="Select Worker:", font=("Arial", 20), background="#f5f5dc").pack(padx=10, pady=10)

        self.carpentry_employee_dropdown = ttk.Combobox(frame_carpentry, state="readonly", font=("Arial", 18), width=40)
        self.carpentry_employee_dropdown.pack(padx=10)
        self.carpentry_employee_dropdown['values'] = ["------"] + [emp[0] for emp in self.carpentry_employees]
        self.carpentry_employee_dropdown.current(0)
        self.carpentry_employee_dropdown.bind("<<ComboboxSelected>>", lambda event: self.handle_employee_selection("Carpentry"))

        # Worker Button (Carpentry)
        tk.Button(frame_carpentry, text="Worker", command=lambda: self.open_employee_management("Carpentry"), 
                bg="#4CAF50", fg="white", font=("Arial", 18, "bold"), width=20).pack(pady=10)  # Green Button

        # Attendance Display
        self.attendance_frame = tk.LabelFrame(self.root, text="Attendance", fg="blue", font=("Arial", 20), bg="#f5f5dc")
        self.attendance_frame.pack(pady=10, fill="x", padx=30)

        # Worker Name Display with increased font size
        self.employee_name_label = tk.Label(self.attendance_frame, text="Worker: None", font=("Arial", 18, "bold"), fg="black", bg="#f5f5dc")
        self.employee_name_label.pack(side="left", padx=20)

        self.attendance_labels = []
        for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]:
            label = tk.Label(self.attendance_frame, text=f"{day}: Pending", fg="gray", font=("Arial", 18), bg="#f5f5dc")
            label.pack(side="left", padx=10)
            self.attendance_labels.append(label)

        # Cash Advance Section
        cash_advance_frame = tk.Frame(self.root, bg="#f5f5dc")
        cash_advance_frame.pack(anchor="nw", padx=20, pady=0)

        self.cash_advance_label = tk.Label(cash_advance_frame, text="Cash Advance: None", font=("Arial", 18, "bold"), bg="#f5f5dc")
        self.cash_advance_label.pack(side="top", anchor="w", pady=0, padx=25)

        tk.Button(cash_advance_frame, text="Update Cash Advance", command=self.update_cash_advance, 
                bg="purple", fg="white", font=("Arial", 17), width=22).pack(side="top", anchor="w", pady=(5, 10))

        # Center Frame for Buttons & Tables
        center_frame = tk.Frame(self.root, bg="#f5f5dc")
        center_frame.pack(pady=0)

        # Button Frames
        button_frame_top = tk.Frame(center_frame, bg="#f5f5dc")
        button_frame_top.pack(pady=0)

        # First Row of Buttons
        button_width = 18
        tk.Button(button_frame_top, text="Track Attendance", command=self.track_attendance, 
                bg="blue", fg="white", font=("Arial", 17), width=button_width).pack(side="left", padx=5)

        tk.Button(button_frame_top, text="Calculate Salary", command=self.calculate_salary, 
                bg="green", fg="white", font=("Arial", 17), width=button_width).pack(side="left", padx=5)

        # Second Row of Buttons
        button_frame_bottom = tk.Frame(center_frame, bg="#f5f5dc")
        button_frame_bottom.pack(pady=10)

        tk.Button(button_frame_bottom, text="Reset Records", command=self.reset_attendance_and_cash_advance, 
                bg="red", fg="white", font=("Arial", 17), width=button_width).pack(side="left", padx=5)

        # Split Tables for Finishing and Carpentry
        tables_container = tk.Frame(center_frame, bg="#f5f5dc")
        tables_container.pack(fill="both", expand=True, padx=30, pady=20)

        # Finishing Table
        table_frame1 = tk.Frame(tables_container, bg="#f5f5dc")
        table_frame1.pack(side="left", fill="both", expand=True, padx=30)

        scrollbar1 = ttk.Scrollbar(table_frame1, orient="vertical")
        scrollbar1.pack(side="right", fill="y")

        # Styling Tables
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("TkDefaultFont", 15))
        style.configure("Treeview", font=("TkDefaultFont", 14))

        self.finishing_tree = ttk.Treeview(table_frame1, columns=("Worker", "Work Status", "Salary", "Cash Advance"), 
                                        show='headings', yscrollcommand=scrollbar1.set)
        for col in self.finishing_tree['columns']:
            self.finishing_tree.heading(col, text=col, anchor='center')
            self.finishing_tree.column(col, anchor='center', width=150)

        self.finishing_tree.pack(side="left", fill="both", expand=True)
        scrollbar1.config(command=self.finishing_tree.yview)

        # Carpentry Table
        table_frame2 = tk.Frame(tables_container, bg="#f5f5dc")
        table_frame2.pack(side="left", fill="both", expand=True, padx=30)

        scrollbar2 = ttk.Scrollbar(table_frame2, orient="vertical")
        scrollbar2.pack(side="right", fill="y")

        self.carpentry_tree = ttk.Treeview(table_frame2, columns=("Worker", "Work Status", "Salary", "Cash Advance"), 
                                        show='headings', yscrollcommand=scrollbar2.set)
        for col in self.carpentry_tree['columns']:
            self.carpentry_tree.heading(col, text=col, anchor='center')
            self.carpentry_tree.column(col, anchor='center', width=150)

        self.carpentry_tree.pack(side="left", fill="both", expand=True)
        scrollbar2.config(command=self.carpentry_tree.yview)

        # Export Data Buttons
        export_frame = tk.Frame(center_frame, bg="#f5f5dc")
        export_frame.pack(pady=15)

        tk.Button(export_frame, text="Export Finishing Data", command=lambda: self.export_to_excel("Finishing"),
                bg="green", fg="white", font=("Arial", 14), width=20).pack(side="left", padx=10)

        tk.Button(export_frame, text="Export Carpentry Data", command=lambda: self.export_to_excel("Carpentry"),
                bg="blue", fg="white", font=("Arial", 14), width=20).pack(side="left", padx=10)

    def handle_employee_selection(self, category):
        """Clears the other dropdown and updates attendance based on selected worker"""

        if category == "Finishing":
            self.carpentry_employee_dropdown.set("------")
        elif category == "Carpentry":
            self.finishing_employee_dropdown.set("------")

        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()

        if worker == "------":
            self.employee_name_label.config(text="Worker: None")
            for label in self.attendance_labels:
                label.config(text=f"{label.cget('text').split(':')[0]}: Pending", fg="gray")
            self.cash_advance_label.config(text="Cash Advance: None")
            return

        # Update attendance display for the selected worker
        self.employee_name_label.config(text=f"Worker: {worker}")

        # Reset all labels to "Pending" before updating
        for label in self.attendance_labels:
            label.config(text=f"{label.cget('text').split(':')[0]}: Pending", fg="gray")

        # ✅ Handle cases where old records use "Employee" instead of "Worker"
        for record in self.data:
            worker_key = "Worker" if "Worker" in record else "Employee" if "Employee" in record else None  
            if worker_key and record[worker_key] == worker and "Attendance" in record:
                attendance = record["Attendance"]
                for day, status in attendance.items():
                    for label in self.attendance_labels:
                        if day in label.cget('text'):
                            label.config(
                                text=f"{day}: {status}",
                                fg="green" if status == "Full" else "orange" if status == "Half" else "red"
                            )
                            break

        self.update_cash_advance_display()

    def update_cash_advance_display(self):
        """Updates the Cash Advance display based on selected worker."""
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()
        
        if worker == "------":
            self.cash_advance_label.config(text="Cash Advance: None")
            return
        
        # Get cash advance amount
        cash_advance = next((emp[2] for emp in self.finishing_employees if emp[0] == worker), None)
        if cash_advance is None:
            cash_advance = next((emp[2] for emp in self.carpentry_employees if emp[0] == worker), None)
        
        # Ensure a default value if cash advance is not found
        if cash_advance is None:
            cash_advance = 0

        # Get salary for reference (if needed)
        salary = next((emp[1] for emp in self.finishing_employees if emp[0] == worker), None)
        if salary is None:
            salary = next((emp[1] for emp in self.carpentry_employees if emp[0] == worker), None)
        
        # Update UI Labels
        self.cash_advance_label.config(text=f"Cash Advance: {cash_advance}")
        self.employee_name_label.config(text=f"Worker: {worker}")  # add Salary: {salary} if need a salary labels

        # Save data after updating the display
        self.save_data(silent=True)
        # self.export_to_excel()  # Also save to Excel and TXT when updating incentive

    def update_cash_advance(self):
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()
        if worker == "------":
            self.root.attributes("-topmost", True)
            messagebox.showerror("Error", "Please select a worker.", parent=self.root)
            self.root.attributes("-topmost", False)
            return

        current_cash_advance = next((emp[2] for emp in self.finishing_employees if emp[0] == worker), None)
        if current_cash_advance is None:
            current_cash_advance = next((emp[2] for emp in self.carpentry_employees if emp[0] == worker), None)

        # Create a bigger window
        cash_advance_window = tk.Toplevel(self.root)
        cash_advance_window.title("Update Cash Advance")
        cash_advance_window.configure(bg="#efebe2")  # Match background color

        # Set window size
        window_width = 400
        window_height = 150

        # Get screen width and height
        screen_width = cash_advance_window.winfo_screenwidth()
        screen_height = cash_advance_window.winfo_screenheight()

        # Calculate position to center the window
        x_position = (screen_width // 2) - (window_width // 2)
        y_position = (screen_height // 2) - (window_height // 2)

        cash_advance_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Make sure it's always on top
        cash_advance_window.lift()
        cash_advance_window.attributes("-topmost", True)

        tk.Label(cash_advance_window, text=f"Enter new cash advance for {worker}:", font=("Arial", 14)).pack(pady=10)
        entry = tk.Entry(cash_advance_window, font=("Arial", 14))
        entry.insert(0, str(current_cash_advance))
        entry.pack(pady=5)

        def save_cash_advance():
            try:
                new_cash_advance = float(entry.get())
                if worker in [emp[0] for emp in self.finishing_employees]:
                    for i, emp in enumerate(self.finishing_employees):
                        if emp[0] == worker:
                            self.finishing_employees[i] = (emp[0], emp[1], new_cash_advance)
                            break
                elif worker in [emp[0] for emp in self.carpentry_employees]:
                    for i, emp in enumerate(self.carpentry_employees):
                        if emp[0] == worker:
                            self.carpentry_employees[i] = (emp[0], emp[1], new_cash_advance)
                            break
                self.update_tree_cash_advance(worker, new_cash_advance)
                self.update_cash_advance_display()
                self.save_data(silent=True)
                cash_advance_window.destroy()  # Close the window after saving
            except ValueError:
                self.root.attributes("-topmost", True)
                messagebox.showerror("Error", "Please enter a valid number.", parent=self.root)
                self.root.attributes("-topmost", False)

        tk.Button(cash_advance_window, text="Save", command=save_cash_advance, font=("Arial", 14), bg="green", fg="white").pack(pady=10)


    def update_tree_cash_advance(self, worker, new_cash_advance):
        for tree in (self.finishing_tree, self.carpentry_tree):
            for item in tree.get_children():
                if tree.item(item, "values")[0] == worker:
                    values = tree.item(item, "values")
                    tree.item(item, values=(values[0], values[1], values[2], values[3], new_cash_advance))

    def open_employee_management(self, category):
        employee_window = tk.Toplevel(self.root)
        employee_window.title(f"Manage {category} Employees")
        employee_window.attributes("-topmost", True)
        self.center_window(400, 550, employee_window)  # Adjusted height for better spacing
        employee_window.configure(bg="#efebe2")  # Match background color

        # Listbox to display employees
        self.employee_listbox = tk.Listbox(employee_window, height=10, width=50, font=("Arial", 14))
        self.employee_listbox.pack(pady=10)

        # Populate listbox based on category
        employees = self.finishing_employees if category == "Finishing" else self.carpentry_employees
        self.update_employee_listbox(employees)

        # Button Frame (Centers buttons)
        button_frame = tk.Frame(employee_window, bg="#efebe2")
        button_frame.pack(pady=10)

        # Button Styling
        button_width = 20
        button_height = 2
        font_style = ("Arial", 16, "bold")

        # Add Employee Button
        tk.Button(button_frame, text="➕ Add", bg="#008000", fg="white", font=font_style,
                width=button_width, height=button_height, relief="raised", bd=3,
                command=lambda: self.add_employee(category)).pack(pady=5)

        # Edit Employee Button
        tk.Button(button_frame, text="✏️ Edit", bg="#0057D9", fg="white", font=font_style,
                width=button_width, height=button_height, relief="raised", bd=3,
                command=lambda: self.edit_employee(category)).pack(pady=5)

        # Delete Employee Button
        tk.Button(button_frame, text="❌ Delete", bg="#D90000", fg="white", font=font_style,
                width=button_width, height=button_height, relief="raised", bd=3,
                command=lambda: self.delete_employee(category)).pack(pady=5)

    def update_employee_listbox(self, workers):
        self.employee_listbox.delete(0, tk.END)
        for worker in workers:
            self.employee_listbox.insert(tk.END, f"{worker[0]} - Salary: {worker[1]}, Cash Advance: {worker[2]}")

    def add_employee(self, category):
        """Adds a new worker without asking for an initial cash advance."""

        # Create a new pop-up window
        add_window = tk.Toplevel(self.root)
        add_window.title("Add Worker")
        add_window.attributes("-topmost", True)
        self.center_window(350, 200, add_window)

        tk.Label(add_window, text="Enter Worker Name:", font=("Arial", 14)).pack(pady=5)
        name_entry = tk.Entry(add_window, font=("Arial", 14), width=30)
        name_entry.pack(pady=5)

        tk.Label(add_window, text="Enter Salary:", font=("Arial", 14)).pack(pady=5)
        salary_entry = tk.Entry(add_window, font=("Arial", 14), width=30)
        salary_entry.pack(pady=5)

        def submit_employee():
            name = name_entry.get().strip()
            salary = salary_entry.get().strip()

            if not name:
                self.root.attributes("-topmost", True)
                messagebox.showerror("Error", "Worker name cannot be empty.", parent=self.root)
                self.root.attributes("-topmost", False)
                return
            try:
                salary = float(salary)
                if salary <= 0:
                    raise ValueError
            except ValueError:
                self.root.attributes("-topmost", True)
                messagebox.showerror("Error", "Please enter a valid positive salary.")
                self.root.attributes("-topmost", False)
                return
            
            cash_advance = 0

            if category == "Finishing":
                self.finishing_employees.append((name, salary, cash_advance))
                self.data.append({"Worker": name, "Attendance": {day: "Pending" for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]}})
            elif category == "Carpentry":
                self.carpentry_employees.append((name, salary, cash_advance))
                self.data.append({"Worker": name, "Attendance": {day: "Pending" for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]}})

            self.update_employee_listbox(self.finishing_employees if category == "Finishing" else self.carpentry_employees)
            self.save_data(silent=True)
            self.update_attendance_display()  # Update attendance display for the new worker
            self.update_dropdowns()

            add_window.destroy()

        # Add a submit button to confirm
        tk.Button(add_window, text="Add Worker", command=submit_employee, font=("Arial", 14), bg="green", fg="white").pack(pady=10)

    def edit_employee(self, category):
        """Edits a worker using a single pop-up window for name & salary"""
        selected_employee = self.employee_listbox.curselection()
        
        if not selected_employee:
            self.root.attributes("-topmost", True)
            messagebox.showerror("Error", "Please select a worker to edit.", parent=self.root)
            self.root.attributes("-topmost", False)
            return
        
        # Extract current details
        employee_data = self.employee_listbox.get(selected_employee[0])
        name, details = employee_data.split(" - ")
        salary_str, cash_advance_str = details.split(", Cash Advance: ")
        salary = float(salary_str.split(": ")[1])
        cash_advance = float(cash_advance_str)

        # Create a pop-up window
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Worker")
        edit_window.attributes("-topmost", True)
        self.center_window(350, 220, edit_window)

        tk.Label(edit_window, text="Edit Worker Name:", font=("Arial", 14)).pack(pady=5)
        name_entry = tk.Entry(edit_window, font=("Arial", 14), width=30)
        name_entry.insert(0, name)
        name_entry.pack(pady=5)

        tk.Label(edit_window, text="Edit Salary:", font=("Arial", 14)).pack(pady=5)
        salary_entry = tk.Entry(edit_window, font=("Arial", 14), width=30)
        salary_entry.insert(0, str(salary))
        salary_entry.pack(pady=5)

        def submit_edit():
            new_name = name_entry.get().strip()
            new_salary = salary_entry.get().strip()

            if not new_name:
                self.root.attributes("-topmost", True)
                messagebox.showerror("Error", "Worker name cannot be empty.", parent=self.root)
                self.root.attributes("-topmost", False)
                return
            try:
                new_salary = float(new_salary)
                if new_salary <= 0:
                    raise ValueError
            except ValueError:
                self.root.attributes("-topmost", True)
                messagebox.showerror("Error", "Please enter a valid positive salary.", parent=self.root)
                self.root.attributes("-topmost", False)
                return
            
            # Update worker details
            if category == "Finishing":
                self.finishing_employees[selected_employee[0]] = (new_name, new_salary, cash_advance)
            elif category == "Carpentry":
                self.carpentry_employees[selected_employee[0]] = (new_name, new_salary, cash_advance)

            self.update_employee_listbox(self.finishing_employees if category == "Finishing" else self.carpentry_employees)
            self.save_data(silent=True)
            self.update_dropdowns()

            edit_window.destroy()

        # Add a submit button to confirm
        tk.Button(edit_window, text="Save Changes", command=submit_edit, font=("Arial", 14), bg="blue", fg="white").pack(pady=10)

    def delete_employee(self, category):
        selected_employee = self.employee_listbox.curselection()
        if selected_employee:
            employee_name = self.employee_listbox.get(selected_employee[0]).split(" - ")[0]

            self.root.attributes("-topmost", True)
            confirm =  messagebox.askyesno("Delete", f"Are you sure you want to delete {employee_name}?", parent=self.root)
            self.root.attributes("-topmost", False)

            if confirm:
                if category == "Finishing":
                    employee_list = self.finishing_employees
                elif category == "Carpentry":
                    employee_list = self.carpentry_employees
                else:
                    return
                
                index_to_delete = next((i for i, emp in enumerate(employee_list) if emp[0] == employee_name), None)

                if index_to_delete is not None:
                    employee_list.pop(index_to_delete)

                self.update_employee_listbox(employee_list)
                self.remove_employee_data(employee_name)  # Remove data from all associated files
                self.save_data(silent=True)  # Save data after deletion
                self.update_dropdowns()

    def remove_employee_data(self, employee_name):
        # Remove attendance and cash advance data for the deleted worker
        self.data = [
            record for record in self.data
            if record.get("Worker", record.get("Employee")).strip().lower() != employee_name.strip().lower()]
        self.save_data(silent=True)  # Save changes to the JSON file

    def track_attendance(self):
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()

        if worker == "------":
            messagebox.showerror("Error", "Please select a worker.", parent=self.root)
            return
        
        if hasattr(self, "attendance_window") and self.attendance_window.winfo_exists():
            self.attendance_window.lift()
            return

        self.attendance_window = tk.Toplevel(self.root)
        self.attendance_window.configure(bg="#efebe2")  # Match background color
        self.attendance_window.title(f"Attendance for {worker}")
        self.attendance_window.attributes("-topmost", True)
        self.center_window(800, 650, self.attendance_window)

        font_style = ("Arial", 14, "bold")
        button_width = 12
        button_height = 2

        button_colors = {"Full": "green", "Half": "orange", "Absent": "red"}
        darkened_colors = {"Full": "#003200", "Half": "#8B4500", "Absent": "#5B0000"}

        frame = tk.Frame(self.attendance_window, bg="#efebe2")
        frame.pack(pady=10, padx=20, fill="both", expand=True)

        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

        # Store button references for updatiung colors
        if not hasattr(self, "attendance_buttons"):
            self.attendance_buttons = {}
        
        # Keep existing attendance if available, otherwise set "Pending"
        existing_record = next((record for record in self.data if record["Worker"] == worker), None)
        if not hasattr(self, "attendance_vars"):
            self.attendance_vars = {}

        for day in days:
            self.attendance_vars.setdefault(day, tk.StringVar(value="Pending"))
            if existing_record and day in existing_record["Attendance"]:
                self.attendance_vars[day].set(existing_record["Attendance"][day])

        for day in days:
            row_frame = tk.Frame(frame, bg="#efebe2")
            row_frame.pack(fill="x", pady=8)

            # Day Label
            tk.Label(row_frame, text=day, font=font_style, width=15, anchor="w").pack(side="left", padx=10, fill="x", expand=True)

            self.attendance_buttons[day] = {}

            # Attendance buttons
            for status in ["Full", "Half", "Absent"]:
                button = tk.Button(row_frame, text=status, bg=button_colors[status], fg="white",
                                   font=font_style, width=button_width, height=button_height,
                                   command=lambda d=day, s=status: self.set_attendance(d, s))
                button.pack(side="left", padx=8, fill="x", expand=True)

                self.attendance_buttons[day][status] = button

                if self.attendance_vars[day].get() == status:
                    button.config(bg=darkened_colors[status])
        
        tk.Button(self.attendance_window, text="Save Attendance", command=lambda: self.submit_attendance(worker),
                  font=font_style, bg="blue", fg="white", width=25, height=2).pack(pady=20)


    def set_attendance(self, day, status):
        self.attendance_vars[day].set(status)

    def submit_attendance(self, worker):
        """Saves attendance and prevents reset when reopening"""
        attendance_record = {day: var.get() for day, var in self.attendance_vars.items()}
        
        # Check if the worker already has attendance recorded
        existing_record = next((record for record in self.data if record["Worker"] == worker), None)
        if existing_record:
            existing_record["Attendance"].update(attendance_record)  # Update existing record
        else:
            self.data.append({"Worker": worker, "Attendance": attendance_record})  # Add new record

        self.update_attendance_display()  # Update attendance display for the selected worker
        self.attendance_window.destroy()
        self.save_data(silent=True)

    def update_attendance_display(self, event=None):
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()
        
        if worker == "------":
            self.employee_name_label.config(text="Worker: None")
            for label in self.attendance_labels:
                label.config(text=f"{label.cget('text').split(':')[0]}: Pending", fg="gray")
            self.cash_advance_label.config(text="Cash Advance: None")
            return

        self.employee_name_label.config(text=f"Worker: {worker}")
        
        # Initialize all labels to "Pending"
        for label in self.attendance_labels:
            label.config(text=f"{label.cget('text').split(':')[0]}: Pending", fg="gray")

        # Find the attendance record for the selected worker
        for record in self.data:
            if record["Worker"] == worker and "Attendance" in record:
                attendance = record["Attendance"]
                for day, status in attendance.items():
                    for label in self.attendance_labels:
                        if day in label.cget('text'):
                            label.config(text=f"{day}: {status}", fg="green" if status == "Full" else "orange" if status == "Half" else "red")
                            break

        self.update_cash_advance_display()  # Update cash advance and incentive display

    def calculate_salary(self):
        """Calculate salary, adds cash advance, and updates the table."""

        # Determine which worker is selected
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()
        
        if worker == "------":
            self.root.attributes("-topmost", True)
            messagebox.showerror("Error", "Please select a worker.", parent=self.root)
            self.root.attributes("-topmost", False)
            return

        # Get worker salary
        salary = next((emp[1] for emp in self.finishing_employees if emp[0] == worker), None)
        if salary is None:
            salary = next((emp[1] for emp in self.carpentry_employees if emp[0] == worker), None)

        # Get worker cash advance
        cash_advance = next((emp[2] for emp in self.finishing_employees if emp[0] == worker), None)
        if cash_advance is None:
            cash_advance = next((emp[2] for emp in self.carpentry_employees if emp[0] == worker), None)

        # Ensure default values if not found
        salary = salary if salary is not None else 0
        cash_advance = cash_advance if cash_advance is not None else 0

        # Get attendance data
        attendance = next((record["Attendance"] for record in self.data if record["Worker"] == worker and "Attendance" in record), {})

        # Calculate total salary based on attendance
        total_salary = sum(salary if status == "Full" else salary / 2 if status == "Half" else 0 for status in attendance.values())

        # Add cash advance to the salary
        final_salary = total_salary + cash_advance

        # Insert into the correct table
        # if worker in [emp[0] for emp in self.finishing_employees]:
        #     self.finishing_tree.insert("", "end", values=(worker, "Calculated", final_salary, cash_advance))
        # elif worker in [emp[0] for emp in self.carpentry_employees]:
        #     self.carpentry_tree.insert("", "end", values=(worker, "Calculated", final_salary, cash_advance))
        if worker in [emp[0] for emp in self.finishing_employees]:
            tree = self.finishing_tree
        elif worker in [emp[0] for emp in self.carpentry_employees]:
            tree = self.carpentry_tree
        else:
            return

        children = tree.get_children()

        last_worker = None
        for child in reversed(children):
            last_values = tree.item(child)["values"]
            if last_values and last_values[0]:
                last_worker = last_values[0]
                break

        worker_name = "" if last_worker == worker else worker

        tree.insert("", "end", values=(worker_name, "Calculated", final_salary, cash_advance))

        # Save updated data
        self.save_data(silent=True)

    def reset_attendance_and_cash_advance(self):
        """Resets attendance and cash advance for the selected worker."""
        worker = self.finishing_employee_dropdown.get() if self.finishing_employee_dropdown.get() != "------" else self.carpentry_employee_dropdown.get()

        if worker == "------":
            self.root.attributes("-topmost", True)
            messagebox.showerror("Error", "Please select a worker.", parent=self.root)
            self.root.attributes("-topmost", False)
            return

        if messagebox.askyesno("Confirm Reset", f"Are you sure you want to reset {worker}'s attendance and cash advance?", parent=self.root):
            for record in self.data:
                if record["Worker"] == worker:
                    record["Attendance"] = {day: "Pending" for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]}

                    for i, emp in enumerate(self.finishing_employees):
                        if emp[0] == worker:
                            self.finishing_employees[i] = (emp[0], emp[1], 0)
                            break

                    for i, emp in enumerate(self.carpentry_employees):
                        if emp[0] == worker:
                            self.carpentry_employees[i] = (emp[0], emp[1], 0)
                            break

                    self.root.attributes("-topmost", True)
                    messagebox.showinfo("Reset Successfully", f"Attendance and cash advance for {worker} have been reset.", parent=self.root)
                    self.root.attributes("-topmost", False)
                    self.update_attendance_display()
                    self.save_data(silent=True)  # Save changes to the JSON file
                    break

    def export_to_excel(self, category):
        """Exports salary and attendance data to an Excel file, filtered by category"""

        # Select the correct worker list
        if category == "Finishing":
            workers_list = self.finishing_employees
        elif category == "Carpentry":
            workers_list = self.carpentry_employees
        else:
            self.root.attributes("-topmost", True)
            messagebox.showerror("Export Error", "Invalid category selected.", parent=self.root)
            self.root.attributes("-topmost", False)
            return

        # Filter attendance data based on selected category
        data = []
        for record in self.data:
            if any(emp[0] == record.get("Worker", record.get("Employee")) for emp in workers_list):
                attendance = record["Attendance"]
                salary = next((emp[1] for emp in workers_list if emp[0] == record.get("Worker", record.get("Employee"))), 0)
                cash_advance = next((emp[2] for emp in workers_list if emp[0] == record.get("Worker", record.get("Employee"))), 0)
                total_salary = sum(salary if status == "Full" else salary / 2 if status == "Half" else 0 for status in attendance.values())

                data.append({
                    "Worker": record.get("Worker", record.get("Employee")),
                    **attendance,
                    "Total Salary": total_salary,
                    "Total Cash Advance": cash_advance
                })

        # Check if there is data to export
        if not data:
            self.root.attributes("-topmost", True)
            messagebox.showerror("Export Error", f"No data available for {category} workers.", parent=self.root)
            self.root.attributes("-topmost", False)
            return

        df = pd.DataFrame(data)

        os.makedirs(self.receipts_folder, exist_ok=True)

        # Get current date and time for the filename
        current_time = datetime.now().strftime("%m-%d-%Y %H-%M-%S").replace(":", "-")
        excel_path = os.path.join(self.receipts_folder, f"{category}_attendance_receipt_{current_time}.xlsx")

        # Create a progress pop-up window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Exporting Data")
        progress_window.geometry("450x150")
        progress_window.attributes("-topmost", True)

        window_width = 450
        window_height = 150
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_position = (screen_width // 2) - (window_width // 2)
        y_position = (screen_height // 2) - (window_height // 2)
        progress_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        frame = tk.Frame(progress_window)
        frame.pack(expand=True, fill="both")

        tk.Label(frame, text=f"Exporting {category} data, please wait...", font=("Arial", 14)).pack(pady=10)

        progress = ttk.Progressbar(frame, orient="horizontal", length=250, mode="indeterminate")
        progress.pack(pady=10)

        progress.start()

        # ✅ Prevents UI freeze by delaying Excel export
        self.root.after(100, lambda: self.save_to_excel(df, excel_path, progress, progress_window, category))

    def save_to_excel(self, df, excel_path, progress, progress_window, category):
        """Saves the DataFrame to Excel without freezing the UI."""
        df.to_excel(excel_path, index=False)
        progress.stop()
        progress_window.destroy()

        # ✅ Export to text file after Excel file is confirmed saved
        self.export_to_text_file(category)

        # Open folder after exporting
        self.open_export_folder(self.receipts_folder)


    def export_to_text_file(self, category_name):
        """Exports total salary and cash advance for a specific category (Finishing or Carpentry)."""

        if category_name.lower() == "carpentry":
            employees = self.carpentry_employees
        elif category_name.lower() == "finishing":
            employees = self.finishing_employees
        else:
            print(f"ERROR: Unknown category '{category_name}'")
            return

        # ✅ Calculate total salary and cash advance
        total_salary = sum(emp[1] for emp in employees)
        total_cash_advance = sum(emp[2] for emp in employees)

        # ✅ Get current date and time for the filename
        current_time = datetime.now().strftime("%m-%d-%Y %H-%M-%S")

        # ✅ Format the receipt
        receipt_text = (
            f"Category: {category_name}\n"
            f"Total Salary: {total_salary}\n"
            f"Total Cash Advance: {total_cash_advance}\n"
        )

        # ✅ Save to file
        text_file_path = os.path.join(self.receipts_folder, f"{category_name}_totals_{current_time}.txt")
        with open(text_file_path, "w") as text_file:
            text_file.write(receipt_text)

    def save_data(self, silent=False):
        """Saves all employee and attendance data to JSON in SCfiles - Workers."""
        os.makedirs(self.base_folder, exist_ok=True)

        data_to_save = {
            "finishing_employees": self.finishing_employees,
            "carpentry_employees": self.carpentry_employees,
            "data": self.data
        }

        with open(self.json_file_path, "w") as file:
            json.dump(data_to_save, file, indent=4)

        # if not silent:  
        #     messagebox.showinfo("Backup Complete", f"Data has been successfully saved!\nLocation: {self.receipts_folder}")


    def load_data(self):
        """Loads salary and attendance data from JSON in SCfiles - Workers."""
        try:
            with open(self.json_file_path, "r") as file:
                data = json.load(file)
                self.finishing_employees = data.get("finishing_employees", [])
                self.carpentry_employees = data.get("carpentry_employees", [])
                self.data = data.get("data", [])

                # ✅ Convert old "Employee" keys to "Worker" (if needed)
                for record in self.data:
                    if "Employee" in record and "Worker" not in record:
                        record["Worker"] = record.pop("Employee")
            
        except FileNotFoundError:
            pass

    def open_export_folder(self, folder_path):
        """Opens the dolder where files are exported."""
        try:
            if os.name == "nt": # Windows
                os.startfile(folder_path)
            elif os.name == "posix": # macOS & Linux
                subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            self.root.attributes("-topmost", True)
            messagebox.showerror("Error", f"Could not open folder: {e}")
            self.root.attributes("-topmost", False)

    def update_dropdowns(self):
        """Dynamically updates the Finishing and Carpentry dropdowns."""

        # Get updated worker lists
        finishing_names = ["------"] + [emp[0] for emp in self.finishing_employees]
        carpentry_names = ["------"] + [emp[0] for emp in self.carpentry_employees]

        # Update dropdown values
        self.finishing_employee_dropdown["values"] = finishing_names
        self.carpentry_employee_dropdown["values"] = carpentry_names

        # Reset selection to "------"
        self.finishing_employee_dropdown.set("------")
        self.carpentry_employee_dropdown.set("------")

        # Update UI to reflect changes
        self.update_attendance_display()

if __name__ == "__main__":
    root = tk.Tk()
    app = SalaryCalculator(root)
    root.mainloop()