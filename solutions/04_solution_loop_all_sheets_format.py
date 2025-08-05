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
Format every sheet in the workbook.
1. Load sales_data.xlsx.
2. Loop through each worksheet.
3. Set the width of column A to 20.
4. Make the header row (row 1) bold.
5. Save the updated workbook as sales_data_formatted.xlsx.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import openpyxl
from openpyxl.styles import Font

workbook = openpyxl.load_workbook('sales_data.xlsx')
for sheet in workbook.worksheets:
    sheet.column_dimensions['A'].width = 20
    for cell in sheet[1]:
        cell.font = Font(bold=True)

workbook.save('sales_data_formatted.xlsx')
