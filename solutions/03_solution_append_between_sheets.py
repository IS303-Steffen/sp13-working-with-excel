'''
OPTIONAL AI GUIDANCE PROMPT
---------------------------
I am a student in an introductory Python class. I am learning many coding
principles for the very first time. I am going to paste in the instructions
to a practice problem that my professor gave me to try before class.
Please be my kind tutor and walk me through how to solve the problem step
by step.

Don't just give me the full solution all at once (unless I later ask for
it). Instead, help me work through it gradually, with clear explanations
and small, easy-to-understand examples. Please use everyday language and
explain things in a simple, friendly way.

INSTRUCTIONS:
-------------
This task appends rows from one workbook to another.
1. Load sales_data.xlsx and select the sheet named Mar.
2. Load summary.xlsx and select the sheet named Q1_Summary.
3. Starting at row 2 of Mar, loop until the Item column (column A)
   is blank. Append each full row (columns A-C) to the end of the
   Q1_Summary sheet.
4. Save your changes to a new file called summary_updated.xlsx.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import openpyxl

source_wb = openpyxl.load_workbook('sales_data.xlsx')
mar_sheet = source_wb['Mar']

dest_wb = openpyxl.load_workbook('summary.xlsx')
summary_sheet = dest_wb['Q1_Summary']

row = 2
while mar_sheet[f'A{row}'].value is not None:
    item = mar_sheet[f'A{row}'].value
    units = mar_sheet[f'B{row}'].value
    revenue = mar_sheet[f'C{row}'].value
    summary_sheet.append([item, units, revenue])
    row += 1

dest_wb.save('summary_updated.xlsx')
