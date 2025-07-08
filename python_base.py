import pandas as pd

# Load all sheets using specific columns (third column is Integer)
df1 = pd.read_excel('test.xlsx', sheet_name=0, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8], dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
df2 = pd.read_excel('test.xlsx', sheet_name=1, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8], dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
df3 = pd.read_excel('test.xlsx', sheet_name=2, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8], dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)
df4 = pd.read_excel('test.xlsx', sheet_name=3, skiprows=2, usecols=[0, 1, 2, 3, 4, 5, 6, 8], dtype={0: str, 1: str, 2: int, 3: str, 4: str, 5: str, 6: str, 8: str}, header=None)

# Insert row number (ID) for the first sheet
df1.insert(0, 'Row Number', range(1, len(df1) + 1))

# Continue row numbering (ID) for the second sheet
next_number1 = len(df1) + 1
df2.insert(0, 'Row Number', range(next_number1, next_number1 + len(df2)))

# Continue row numbering (ID) for the third sheet
next_number2 = len(df1) + len(df2) + 1
df3.insert(0, 'Row Number', range(next_number2, next_number2 + len(df3)))

# Continue row numbering (ID) for the fourth sheet
next_number3 = len(df1) + len(df2) + len(df3) + 1
df4.insert(0, 'Row Number', range(next_number3, next_number3 + len(df4)))

# Concatenate all sheets (rows) into a single dataframe
result_df = pd.concat([df1, df2, df3, df4], axis=0)

# Save the result
result_df.to_csv('modified_file.csv', sep=';', index=False, header=False, encoding='utf-8')