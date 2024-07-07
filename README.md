# VBA-report-splitter
This code was created to split a large report into smaller pieces because the original report could not be uploaded to a certain system. The output files will be generated in the same folder as the selected source file. The source file will be split every 350 rows. For example, if the source file contains 2500 rows, there will be 8 output files in the end.

# How to use it
1. Put this code in a Macro-enabled Excel file.
2. Create a button and assign the btnSelectFile_Click() macro to it.
3. Modify the code according to the specific requirements of your file (e.g., worksheet names to delete, the row where the copy process should start, etc.).
