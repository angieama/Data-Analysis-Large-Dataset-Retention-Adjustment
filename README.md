# Data-Analysis-Large-Dataset-Retention-Adjustment
This VBA macro updates "Retention" values in Excel files across subfolders. It prompts for a new value, searches column C for "Retention", and updates the adjacent cell in column D if numeric. Handles merged cells, logs changes, and saves originals. Optimized for performance with disabled updates. Ideal for bulk data.
Data Analysis: Large Dataset Retention Adjustment
Overview
This VBA macro automates updating "Retention" values in Excel files across subfolders. Designed for large datasets (e.g., millions of records), it targets column D adjacent to "Retention" in column C, handling merged cells and skipping non-numeric/empty ones. Processes .xlsx, .xlsm, .xls files in place without backups—use with caution.
Features

Prompts for a single new retention value (numeric).
Searches subfolders in a specified main path (e.g., C:\Desktop\Files\2021\4million\).
Performance optimizations: Disables screen updates, alerts, and auto-calculation.
Logs details to Immediate Window; summarizes changes, skips, and files processed.
Error handling for locked/inaccessible files.

Requirements

Microsoft Excel with VBA enabled.
Files must have consistent structure: "Retention" in column C, value in D (possibly merged with E/F).

Usage

Open Excel and press Alt+F11 to access VBA Editor.
Insert a new module and paste the code from UpdateRetentionValuesInSubfolders.bas.
Modify mainInputFolderPath and offsetColumns as needed.
Run the macro: Tools > Macros > UpdateRetentionValuesInSubfolders.
Enter new value and confirm modification.

Warning: Modifies originals! Back up files first.
Code Structure

Input Handling: User prompt and validation.
File System Loop: Iterates subfolders and files using FileSystemObject.
Worksheet Processing: Finds "Retention", updates target cell.
Cleanup: Restores Excel settings, displays summary.

Limitations

Skips non-Excel files or unsupported extensions.
Assumes no password protection.
Optional: Uncomment to overwrite non-numeric values.

For issues, check Immediate Window (Ctrl+G).
Last updated: October 19, 2025
