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
This exercise writes totals to a new sheet.
1. Load sales_data.xlsx.
2. For each sheet Jan, Feb, and Mar, sum the numbers in column C
   starting at row 2 and continuing until you reach a blank cell.
3. Create a new sheet called Totals and write the three totals into
   cells A1, A2, and A3.
4. Save your changes as a new workbook named sales_data_totals.xlsx
   so the original file is preserved.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import openpyxl

workbook = openpyxl.load_workbook('sales_data.xlsx')
months = ['Jan', 'Feb', 'Mar']
totals = []
for month in months:
    sheet = workbook[month]
    total = 0
    row = 2
    while sheet[f'C{row}'].value is not None:
        total += sheet[f'C{row}'].value
        row += 1
    totals.append(total)

tot_sheet = workbook.create_sheet('Totals')
for idx, val in enumerate(totals, start=1):
    tot_sheet[f'A{idx}'] = val

workbook.save('sales_data_totals.xlsx')
