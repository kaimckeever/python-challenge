# data-challenge.py

import csv
from datetime import date
from openpyxl import load_workbook

wb = load_workbook(filename="in.xlsx")

total_sales = wb['Total Sales']
per_sales = wb['Percentage of Sales']
lookup = wb['Lookup Table']
for row in per_sales.iter_rows(min_row=2, max_row=321, values_only=True):
    yearwk = date(row[1].year,row[1].month,row[1].day).isocalendar()
    col0 = str(yearwk[0]) + '0' + str(yearwk[1])


# with open('out.csv', 'w', newline='') as csvfile:
    # writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
    # for row in per_sales.iter_rows(min_row=2, max_row=7, values_only=True):
    #     yearwk = date(row[1].year,row[1].month,row[1].day).isocalendar()
    #     col0 = str(yearwk[0]) + '0' + str(yearwk[1])