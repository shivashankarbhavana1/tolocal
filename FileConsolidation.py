import pandas as pd
import config


def file_consolidation(file_path):
    # Define the file path
    #file_path = r"C:\Users\Shivashankar Bhavana\OneDrive - The Boston Consulting Group, Inc\Documents\LEARNING\DF\Hydrogen\Dashboard\Data for dashboard\Demand model data\Global Demand Model - Master File - v97.xlsx"

    # Load the specific sheet from the Excel file without assumptions about the header
    df = pd.read_excel(file_path, sheet_name='Consolidated hydrogen demand', header=None)

    # Find the index of the row containing the header 'Geography'
    header_row_idx = df.index[df.apply(lambda x: 'Geography' in x.values, axis=1)][0]

    # Set the new header based on the found index and remove rows above the header row
    df = pd.read_excel(file_path, sheet_name='Consolidated hydrogen demand', header=header_row_idx)

    # After setting the headers, drop any columns if their entire content is NaN
    df = df.dropna(axis=1, how='all')

    # Rename the 'Fuel' column to 'Fuel split.Fuel'
    df.rename(columns={'Fuel': 'Fuel split.Fuel'}, inplace=True)

    # Concatenate fields to create 'Fuel split.Helper'
    df['Fuel split.Helper'] = df['Geography'].astype(str) + df['Application'].astype(str) + df['Fuel split.Fuel'].astype(str)

    # Convert all column headers to strings to ensure operations are valid
    df.columns = df.columns.map(str)

    # Identify columns that are numeric and exactly four characters long (typical year format)
    year_columns = [col for col in df.columns if col.isdigit() and len(col) == 4]
   # print("Year columns identified based on headers:", year_columns)

    # Append 'C' to identified year columns
    rename_dict = {year: f"{year}C" for year in year_columns}
    df.rename(columns=rename_dict, inplace=True)

    # Print updated column headers
    #print("Updated column headers:", df.columns)
    df = df[['Fuel split.Helper'] + [col for col in df.columns if col != 'Fuel split.Helper']]

    print("File consolidation is complete")

    # Format the output path using the extracted user name
    output_path = fr"C:\Users\{config.user_name}\The Boston Consulting Group, Inc\PDT - Hydrogen Models\Automated data for dashboard\Demand Model Data.xlsx"

    #output_path = r"C:\Users\Shivashankar Bhavana\The Boston Consulting Group, Inc\PDT - Hydrogen Models\Automated data for dashboard\Demand Model Data.xlsx"

    with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, index=False, sheet_name='Output data')


if __name__ == '__main__':
    file_consolidation()




