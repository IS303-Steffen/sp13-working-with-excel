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
This problem shows how to open an existing workbook and read data.
1. Use openpyxl.load_workbook to open sales_data.xlsx.
2. Select the sheet named Jan.
3. Read the value in cell B2 and print it.
4. Do not modify or save the workbook.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import openpyxl

workbook = openpyxl.load_workbook('sales_data.xlsx')
sheet = workbook['Jan']
value = sheet['B2'].value
print(value)
