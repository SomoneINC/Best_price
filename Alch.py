import pandas as pd
import random
import os
link = "https://www.rimi.lv/e-veikals/lv/produkti/alkoholiskie-dzerieni/c/SH-1?currentPage=1&pageSize=20&query=%3Arelevance%3AallCategories%3ASH-1%3AassortmentStatus%3AinAssortment%3AcategoryNamePath%3AKoktei%25C4%25BCi"
# Create a DataFrame with a link
data = {'Yarn': ["Name"]}
df = pd.DataFrame(data)
random_number = random.randint(100000, 999999)
# Create a Pandas Excel writer using XlsxWriter as the engine
file_name = f"{random_number}.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

# Write the DataFrame to the Excel file
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a hyperlink format to the workbook
url_format = workbook.add_format({'color': 'blue', 'underline': True})

# Add the hyperlink to the cell
worksheet.write_url('B1', link, url_format, string='yes')

# Save the Excel file
writer._save()
os.system(f'start excel {file_name}')