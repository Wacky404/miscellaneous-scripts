import pandas as pd

df = pd.read_excel(r"C:\Users\Wayne\Documents\march112024_customerlist.xlsx")

date = '2023-01-01'
filtered_df = df.query("last_login <= @date")

filtered_df.to_csv(f'inactivesince{date}')


