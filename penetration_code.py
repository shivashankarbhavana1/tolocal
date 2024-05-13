import pandas as pd
import numpy as np
import config

def filterandclean():

    # Define the path to the original Excel file
    #original_file_path = r'C:\Users\Shivashankar Bhavana\OneDrive - The Boston Consulting Group, Inc\Documents\LEARNING\DF\Hydrogen\Dashboard\Data for dashboard\Demand model data\input sheet.xlsx'

    original_file_path = config.file_path

    # Load the Excel file into a DataFrame from Sheet2
    df = pd.read_excel(original_file_path, sheet_name='Penetration assumptions')
    # Remove the first 10 rows from the DataFrame
    df = df.iloc[10:]

    # Remove specific columns indexed as 0, 1, 3
    df = df.drop(df.columns[[0, 1, 3]], axis=1)

    # First, handle 'Unnamed' columns from the 6th column onward before removing suffixes
    df_first_five = df.iloc[:, :5]
    df_after_five = df.iloc[:, 5:]
    df_after_five = df_after_five.loc[:, ~df_after_five.columns.str.contains('^Unnamed')]
    df = pd.concat([df_first_five, df_after_five], axis=1)

    # Function to remove pandas' appended suffixes
    def remove_suffixes(df):
        df.columns = df.columns.str.replace(r'\.\d+$', '', regex=True)
        return df

    # Apply the function to clean up column names
    df = remove_suffixes(df)

    # Continue with the rest of your processing
    # Get all unique elements in column 0, excluding NA/blank
    unique_elements = df.iloc[:, 0].dropna().unique()

    # Combine 'Region' and 'Reference' with the list of unique elements to create a filter list
    filter_list = [x for x in unique_elements if x == x and x.strip() != '']
    filter_list.extend(['Region', 'Reference'])

    # Filter out rows where column 1 contains any value in filter_list (handling case sensitivity and spaces)
    df_filtered = df[~df.iloc[:, 1].str.strip().str.lower().isin([x.lower().strip() for x in filter_list])]

    # Additional step to remove rows with NaN values in column 3 (index 2)
    df_filtered = df_filtered[df_filtered.iloc[:, 2].notna()]

    # Fill blank values in column 0 with the value above on df_filtered
    df_filtered.iloc[:, 0] = df_filtered.iloc[:, 0].ffill()

    # Reset index after filtering
    df_filtered.reset_index(drop=True, inplace=True)

    # Find the index of "Global penetration rates" and print it
    index_global = df_filtered[df_filtered.iloc[:, 1].str.strip().str.lower() == 'global penetration rates'].index
    if not index_global.empty:
        # Remove all rows from this index onwards
        df_filtered = df_filtered.iloc[:index_global[0], :]



    # Define a new filename to save the modified DataFrame
   # new_file_path = r'C:\Users\shivashankar bhavana\Downloads\cleaned_sheet.xlsx'

    # Save the modified DataFrame to the new Excel file
    #df_filtered.to_excel(new_file_path, index=False)

    print("Data processing complete. The file has been saved under a new name.")

    return df_filtered
def process_and_save_data(input_df, output_excel_path):
    # Load the original Excel data
    df_original = rename_duplicates(input_df)

    # Obtain unique elements from the first row
    unique_values = df_original.iloc[0].unique()

    # Filter out non-numeric values and convert numeric values to integers
    clean_years = [int(value) for value in unique_values if
                   np.issubdtype(type(value), np.number) and not np.isnan(value)]
    # Sort the years
    clean_years.sort()


    # Define scenario mappings
    scenario_map = {
        'STEPS': 'Stated policies scenario',
        'NZE': 'Net zero emissions scenario',
        'SDS': 'Sustainable development scenario'
    }

    # Unpivot the data while retaining the year values and including comments
    df_melted = pd.DataFrame()
    scenarios = ['STEPS', 'NZE', 'SDS']
    #years = ['2020', '2025', '2030', '2035', '2040', '2045', '2050']


    for scenario in scenarios:
        for i, year in enumerate(clean_years):
            temp_df = df_original[['Unnamed: 0', 'Unnamed: 1', scenario + ('' if i == 0 else f'.{i}')]].copy()
            temp_df.columns = ['Application', 'Region', 'Percentage']
            temp_df['Scenario'] = scenario_map[scenario]
            formatted_year = f"1/1/{year}"
            temp_df['Year'] = formatted_year

            temp_df['Comments'] = df_original[
                'Comments' if scenario == 'STEPS' else ('Comments.1' if scenario == 'NZE' else 'Comments.2')]
            df_melted = pd.concat([df_melted, temp_df], ignore_index=True)

    # Clean up data by removing rows where 'Application' or 'Region' is NaN
    df_melted_cleaned = df_melted.dropna(subset=['Application', 'Region'])

    # Save the cleaned and transformed data to an Excel file
    #df_melted_cleaned.to_excel(output_excel_path, index=False)

    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_melted_cleaned.to_excel(writer, index=False, sheet_name='Penetration data')

    print("Penetration Data has been processed and saved to Excel")

def rename_duplicates(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [
            dup + '.' + str(i) if i != 0 else dup for i in range(cols[cols == dup].count())
        ]
    df.columns = cols
    return df

# Apply the function to df_filtered
def transformdata():

    # Define your paths
    input_df = filterandclean()
    #output_excel_path = r'C:\Users\shivashankar bhavana\Downloads\final_transformed_output_sheet.xlsx'
    output_excel_path = fr"C:\Users\{config.user_name}\The Boston Consulting Group, Inc\PDT - Hydrogen Models\Automated data for dashboard\Demand Model Data.xlsx"

    # Run the processing function
    process_and_save_data(input_df, output_excel_path)


if __name__ == '__main__':
    transformdata()
