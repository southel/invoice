from mailmerge import MailMerge
from datetime import date, datetime
import sys
import os
import comtypes.client

TEMPLATE = 'sample-template.docx'
DEFAULT_RATE = 20
DEFAULT_HOURS = 40
document = MailMerge(TEMPLATE)
# getting the invoice nubmer
with open('num.txt', 'r+') as number:
    inv_num = number.readline().strip()
    number.seek(0)
    number.write(str(int(inv_num) + 1))
    number.truncate()
print(f'Invoice Number: {inv_num}')
done = False
# get the current date in date-month string-year format
dte = '{:%d-%b-%Y}'.format(date.today())
# incase you want to backdate ask for previous date
inpt = input('Date [YYYY-MM-DD]? (Enter for today) ')
# parse the date without error handling because we like
# to live on the edge
if inpt:
    dte = '{:%d-%b-%Y}'.format(datetime.strptime(inpt, "%Y-%m-%d"))
# the reason this is in a loop is so that you can have multiple
# line items with different rates
# if you also want different descriptions that the user enters
# you can add that here as well, I didn't need that so I didn't
# implement that feature
total_due = 0
inv_rows = []
while not done:
    row = {}
    inv_rows.append(row)
    rate = input(f'Rate? (Enter for {str(DEFAULT_RATE)} )')
    rate = DEFAULT_RATE if not rate else float(rate)
    qty = input(f'Quantity? (Enter for {str(DEFAULT_HOURS)})')
    qty = DEFAULT_HOURS if not qty else float(qty)
    row['rate'] = f"${rate:.2f}"
    row['qty'] = f"{qty:.2f}"
    total = rate * qty
    total_due += total
    row['total'] = f"${total:.2f}"
    done = input('More? [y/n] ') == 'n'

# injecting the data into the template
document.merge(
    date=dte,
    total_due=f"${total_due:.2f}",
    inv_num=inv_num,
    )
document.merge_rows('qty', inv_rows)
# the file name is the current date
word_file = f"{os.getcwd()}\{dte}.docx"
document.write(word_file)

# converting to PDF
word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(word_file)
doc.SaveAs(f"{os.getcwd()}\{dte}.pdf", FileFormat=17)
doc.Close()
word.Quit()
# deleting filled word file if you want to keep the word file
# then remove this line
os.remove(word_file)
