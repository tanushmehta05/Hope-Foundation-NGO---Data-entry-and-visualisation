Description:

This project is a Python-based data entry and visualization tool designed for the HOPE FOUNDATION. The application provides a user-friendly interface for entering, managing, and visualizing student activity data. It supports multiple entries, generates visualizations, and exports data to Excel with embedded graphs.

Features
    User-Friendly Interface: Designed with customtkinter for a dark-themed, intuitive experience.
    Data Entry: Allows entry of various activity fields for multiple students.
    Activity Tracking: Supports multiple activity trackers per student.
    Visualizations: Generates pie charts and bar graphs for activity distribution.
    Data Export: Exports data to Excel with serial/tracker numbers and embedded graphs using openpyxl.
    User Notifications: Provides feedback for successful actions and input warnings.
    Student Count: Displays the total number of students added.

    -------------------------

    Usage
        Run the application:
        Enter student data into the provided fields.
        Click "Add Tracker" to save the current tracker.
        Use "Add New Student" to reset the form for a new student.
        Repeat the steps if u want to add another information
        Visualize data using the "Visualize Pie Chart" and "Visualize Bar Graph" buttons.
        Save all data to Excel by clicking "Save to Excel".
        Exit the application with the "Exit" button.

    -------------------------

    Dependencies
        Python 3.x
        customtkinter
        matplotlib
        pandas
        tkinter
        openpyxl

Acknowledgements
    This tool was developed to support the HOPE FOUNDATION in efficiently managing and visualizing student activity data.