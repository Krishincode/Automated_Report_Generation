import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from datetime import datetime

# 1. Read CSV
df = pd.read_csv('data.csv')

# 2. Clean data
df.drop_duplicates(inplace=True)
df.fillna(0, inplace=True)

# 3. Summary
summary = df.groupby('Category').agg({'Sales':'sum','Quantity':'sum'}).reset_index()

# 4. Chart
plt.figure(figsize=(8,5))
plt.bar(summary['Category'], summary['Sales'], color='skyblue')
plt.title('Sales by Category')
plt.xlabel('Category')
plt.ylabel('Total Sales')
plt.savefig('sales_chart.png')
plt.close()

# 5. Excel report
wb = Workbook()
ws = wb.active
ws.title = "Summary"
for r in dataframe_to_rows(summary, index=False, header=True):
    ws.append(r)

img = Image('sales_chart.png')
img.anchor = 'E5'
ws.add_image(img)

report_file = f'report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
wb.save(report_file)
print(f"Report generated: {report_file}")
