import openpyxl
from datetime import datetime

wb = openpyxl.load_workbook('JAPAN EXPENSE.xlsx', data_only=True)
ws = wb['Japan Trip']

expenses = []
current_date_label = ''
day_counter = 0

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
    date_cell = row[1]
    transaction_cell = row[2]
    total_cell = row[4]
    jes_cell = row[10]
    cha_cell = row[11]
    joyce_cell = row[12]
    joh_cell = row[13]

    if date_cell.value is not None:
        if isinstance(date_cell.value, datetime):
            day_counter += 1
            actual_day = date_cell.value.year % 100
            current_date_label = f'Day {day_counter} (May {actual_day})'
        elif isinstance(date_cell.value, str) and date_cell.value.strip().upper() == 'OTHERS':
            current_date_label = 'Others (Pre-booked)'
            day_counter += 1

    transaction = transaction_cell.value
    if transaction is None:
        continue
    transaction = str(transaction).strip()
    if not transaction:
        continue
    if transaction.startswith('/') and transaction[1:].strip().isdigit():
        continue

    jes = jes_cell.value if isinstance(jes_cell.value, (int, float)) else 0
    cha = cha_cell.value if isinstance(cha_cell.value, (int, float)) else 0
    joyce = joyce_cell.value if isinstance(joyce_cell.value, (int, float)) else 0
    joh = joh_cell.value if isinstance(joh_cell.value, (int, float)) else 0

    # Only include rows that have actual per-person amounts (K/L/M/N)
    if jes == 0 and cha == 0 and joyce == 0 and joh == 0:
        continue

    total = jes + cha + joyce + joh

    expenses.append({'day': current_date_label, 'txn': transaction, 'total': total, 'jes': jes, 'cha': cha, 'joyce': joyce, 'joh': joh})

print(f'Total transactions: {len(expenses)}')
grand_total = sum(e['total'] for e in expenses)
jes_total = sum(e['jes'] for e in expenses)
cha_total = sum(e['cha'] for e in expenses)

print(f'Grand Total (sum of all): {grand_total:.2f}')
print(f'JES total: {jes_total:.2f}')
print(f'CHA total: {cha_total:.2f}')
print(f'JOYCE total: {sum(e["joyce"] for e in expenses):.2f}')
print(f'JOH total: {sum(e["joh"] for e in expenses):.2f}')
print()
print('Excel summary says:')
print(f'  TOTAL ALL - JJ: {ws.cell(9,20).value}, CHA: {ws.cell(9,21).value}')
print(f'  CREDIT CARD total: {ws.cell(12,19).value}')
print(f'  DEBIT CARD total: {ws.cell(15,19).value}')
print()
print('--- All parsed transactions ---')
for e in expenses:
    print(f"  {e['day']:25s} | {e['txn']:40s} | Total: {e['total']:>10.2f} | JES: {e['jes']:>9.2f} | CHA: {e['cha']:>9.2f}")
