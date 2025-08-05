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
Insert a SUM formula after the data.
1. Load sales_data.xlsx and select the sheet named Jan.
2. Starting at row 2, find the last row that contains data in column C.
3. Two rows below the last data row, write a formula in column C that
   sums the range C2:C<last_row>.
4. Save the workbook as sales_data_formula.xlsx.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import openpyxl

workbook = openpyxl.load_workbook('sales_data.xlsx')
sheet = workbook['Jan']

row = 2
while sheet[f'C{row}'].value is not None:
    row += 1

sum_row = row + 2
sheet[f'C{sum_row}'] = f"=SUM(C2:C{row - 1})"

workbook.save('sales_data_formula.xlsx')
