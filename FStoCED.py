from tkinter.filedialog import askdirectory
import pandas as pd
import glob

path = askdirectory()
print(path)

location = path + '/*.xlsx'

# List of Excel files to merge
excel_files = glob.glob(location)

print(excel_files)


# Define a function to merge Excel files into a single DataFrame
def merge_excel_files(file_list):
    # merged_df = pd.DataFrame()
    df_list = []
    # Set first line to header
    fl = 5
    for file in file_list:
        try:
            # Read each Excel file into a DataFrame
            df = pd.read_excel(file, skiprows=fl, usecols="F:AA")

            # Set first line to exclude header
            fl = 5

            df_list.append(df)

            # Append the DataFrame to the merged_df
            # merged_df = merged_df.concat(df, ignore_index=True)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    merged_df = pd.concat(df_list)
    return merged_df

# Define a function to merge Excel files into a single DataFrame
def parse_excel_files(file_list):
    # merged_df = pd.DataFrame()
    df_list = []
    # Set first line to header
    fl = 5
    for file in file_list:
        try:
            # Read each Excel file into a DataFrame
            df = pd.read_excel(file, skiprows=fl, usecols="F:AA")

            # Set first line to exclude header
            fl = 5

            print(df.head(10))

            # Append the DataFrame to the merged_df
            # merged_df = merged_df.concat(df, ignore_index=True)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    # merged_df = pd.concat(df_list)
    return

# Call the merge_excel_files function and store the result
merged_data = merge_excel_files(excel_files)

#  parse_excel_files(excel_files)

print(merged_data)

# determining the name of the file
file_name = 'merged_data.xlsx'

# saving the excel
merged_data.to_excel(file_name)
print('DataFrame is written to Excel File successfully.')

# You can now work with the merged_data DataFrame as needed.

for index, row in merged_data.iterrows():
    fn = row['fullName']
    names = fn.split(";")
    print(row['relationshipToHead'], names)