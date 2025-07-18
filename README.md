# Excel Daily Sheets
**This project is for demonstration purposes only. The code is not open for modification or redistribution.**
## 🔧 Steps to Use the Macro for Creating Daily Appointment Sheets
### 📁 1. Download & Enable Macros
•	Download the Excel file to your local computer.
•	Enable Editing and Enable Macros when prompted.
________________________________________
### 📅 2. Setup Holidays Sheet
•	Go to the sheet named holiday.
•	In the day column, input all public holidays in date format (e.g., 2025/01/01).
•	Then, drag down the following columns to match the full date range:
o	calendar reference
o	first day
o	The column to the right of first day
(These columns are used for calendar display and holiday-based cell formatting.)
•	If data is already present, you can skip this step.
________________________________________
### 🖱️ 3. Run the Macro to Create Sheets
•	Go to the create sheet tab.
•	Click the Create Sheet button (ensure macros are enabled).
•	Follow the prompts:
1.	Enter the name of the original sheet to copy (usually the first day template).
2.	Enter the start date of the appointment period (format: yyyy/mm/dd).
3.	Enter the end date of the appointment period (format: yyyy/mm/dd).
4.	Enter the weekday(s) to be hidden (e.g., 1 for Monday, 7 for Sunday).
	Enter one number at a time and press OK.
	If no more days to hide, enter 0 and press OK.
✅ The macro will:
•	Duplicate the original sheet for each day in the date range.
•	Add the correct date to each new sheet.
•	Automatically hide the sheets that fall on non-operating days (e.g., weekends or holidays).
________________________________________
###  4. Hide Setup Sheets & Save
•	After generation is complete:
o	Right-click on the tabs origin, holiday, and create sheet, and select Hide.
•	Then go to File > Save As:
o	Choose a new file name.
o	Make sure to save as Excel Workbook (.xlsx).
•	Close the file when done.
(Do not overwrite the original file — use it again for other date ranges if needed.)
________________________________________

## 🔐 Worksheet Protection
•	All generated sheets are protected to prevent accidental changes.
•	Only the appointment area and note column are editable.
•	To unlock or modify the protection, use the password: 0000

## License
This project is for demonstration purposes only. The code is not open for modification or redistribution.

