import pandas as pd

#кратко - код просто разбивает товарные предложения внутри ячейки на отдельные строки товара. Обрати внимание - названия столбцов в инпуте переименованы так, чтобы код мог работать.

# Replace 'input.xls' with your XLS file name
input_filename = 'input.xlsx'
output_filename = 'output.xlsx'  # Output XLSX file name

# Load the XLS data into a pandas DataFrame
xls_df = pd.read_excel(input_filename)

# Create empty lists to hold the new rows
new_rows = []

# Iterate through rows in the DataFrame
for index, row in xls_df.iterrows():
    materials = row['Materials'].split('|') if isinstance(row['Materials'], str) else [""]
    prices = row['Prices'].split('|') if isinstance(row['Prices'], str) else [""]
    quantities = row['Quantities'].split('|') if isinstance(row['Quantities'], str) else [""]

    # Find the maximum number of splits among the columns
    num_splits = max(len(materials), len(prices), len(quantities))

    # Create new rows for each split value
    for i in range(num_splits):
        new_row = row.copy()  # Copy the existing row
        new_row['Materials'] = materials[i] if i < len(materials) else ""
        new_row['Prices'] = prices[i] if i < len(prices) else ""
        new_row['Quantities'] = quantities[i] if i < len(quantities) else ""
        new_rows.append(new_row)  # Append the new row to the list

# Create a new DataFrame from the list of new rows
new_df = pd.DataFrame(new_rows)

# Write the new DataFrame to an XLSX file
new_df.to_excel(output_filename, index=False)