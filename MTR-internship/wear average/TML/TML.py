import pandas as pd
import math
import pandas as pd
from tkinter import filedialog
from datetime import datetime
import re

def read_data(raw_data_path, name):
    """
    Reads the selected .xlsx file. Combines the four wire columns into one column and returns the dataframe.

    Args:
        raw_data_path (str): The path to the .xlsx file.
        name (str): The name of the sheet to be returned.

    Returns:
        (DataFrame): The modified reqeusted sheet.
    """
    data = pd.read_excel(raw_data_path, name)
    data.columns = data.iloc[1]
    data = data.iloc[2:].reset_index().drop(columns='index').dropna(how='all')
    data = data.drop(data[data['LINE'].isin(['Date', 'LINE'])].index).reset_index(drop=True)
    return data[['TRACK', 'CHAINAGE', 'RWH1mm', 'RWH2mm', 'RWH3mm', 'RWH4mm']]

def get_individual_report(data, lookup_table):
    """
    Calculates the mean, standard deviation, and percentage of wear of the given data.
    Stores the results in a DataFrame and returns it.

    Args:
        data (DataFrame): The given input data.
        lookup_table (DataFrame): The Tension Length lookup table for respective section and track of the line.

    Returns:
        (DataFrame): A DataFrame that contains the calculated mean, standard deviation, and percentage of wear at each tension length.
    """
    report = pd.DataFrame()
    for index, row in lookup_table.iterrows():
        location = data[(data['CHAINAGE'] >= row['Overlap FromM']) & (data['CHAINAGE'] <= row['Overlap ToM'])].index
        frames = [data.loc[location, 'RWH1mm'], data.loc[location, 'RWH2mm'], data.loc[location, 'RWH3mm'], data.loc[location, 'RWH4mm']]
        new_df = pd.concat(frames)
        mean = new_df.mean()
        sd = new_df.std()
        if mean != 0:
            wear_percentage = ((((math.acos((mean-6.6)/6.6))*(6.6*6.6))-((6.6*(math.sin(math.acos((mean-6.6)/6.6))))*(mean-6.6)))/120)*100
        else:
            wear_percentage = 0
        temp = pd.DataFrame({
            'TML': [row['Tension Length']], 
            'Mean of Remaining': [mean], 
            'Mean-SD of Remaining': [mean - sd], 
            'Mean-2SD of Remaining': [mean - (2 * sd)], 
            'Mean-3SD of Remaining': [mean - (3 * sd)], 
            '% of wear': [wear_percentage]
        })
        report = pd.concat([report, temp])
    return report.reset_index().drop(columns=['index'], axis=1)

def main():
    raw_data_path = filedialog.askopenfilename()
    print("raw:", raw_data_path)
    print(type(raw_data_path))
    print("Reading Files...")

    # ---------- Reads input data ----------
    TML_up = read_data(raw_data_path, 'up')
    TML_down = read_data(raw_data_path, 'down')
    # ---------- Reads input data ----------

    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    # ---------- Loads lookup tables ----------
    TML_lookup_up = pd.read_excel('./TML/TML.xlsx', sheet_name='TML')
    TML_lookup_up = TML_lookup_up[['tml_up_from', 'tml_up_to', 'tml_up_tl']]\
        .rename({'tml_up_from': 'Overlap FromM', 'tml_up_to': 'Overlap ToM', 'tml_up_tl': 'Tension Length'}, axis=1).dropna(how='all')
    
    TML_lookup_down = pd.read_excel('./TML/TML.xlsx', sheet_name='TML')
    TML_lookup_down = TML_lookup_down[['tml_dn_from', 'tml_dn_to', 'tml_dn_tl']]\
        .rename({'tml_dn_from': 'Overlap FromM', 'tml_dn_to': 'Overlap ToM', 'tml_dn_tl': 'Tension Length'}, axis=1).dropna(how='all')
    # ---------- Loads lookup tables ----------

    # Gets reports for each section and track
    TML_up_report = get_individual_report(TML_up, TML_lookup_up)
    TML_down_report = get_individual_report(TML_down, TML_lookup_down)
    # Gets reports for each section and track

    # Combines up and down track into one report
    TML_report = pd.concat([TML_up_report, TML_down_report], ignore_index=True)

    # ---------- Sorts the table by the tension length ----------
    # M01 --> U01 D02 --> K01 --> 01
    ocs = TML_report[TML_report['TML'].str[0] == 'M']
    ocs['id'] = ocs['TML'].str[1:].astype(int)
    ocs = ocs.sort_values(by='id').drop(columns='id')

    scl_etse = TML_report[(TML_report['TML'].str[0] == 'U') | (TML_report['TML'].str[0] == 'D')]
    scl_etse['id'] = scl_etse['TML'].str[1:].astype(int)
    scl_etse = scl_etse.sort_values(by='id').drop(columns='id')

    ksl = TML_report[TML_report['TML'].str[0] == 'K']
    ksl['id'] = ksl['TML'].str[1:].astype(int)
    ksl = ksl.sort_values(by='id').drop(columns='id')

    wrl = TML_report[(TML_report['TML'].str[0] != 'M') &
                     (TML_report['TML'].str[0] != 'U') &
                     (TML_report['TML'].str[0] != 'D') &
                     (TML_report['TML'].str[0] != 'K')]
    wrl = wrl.sort_values(by='TML')

    TML_report = pd.concat([ocs, scl_etse, wrl], ignore_index=True)
    # ---------- Sorts the table by the tension length ----------

    # ---------- Saving output .xlsx files ----------
    date_pattern = re.search(r'\d{6}', raw_data_path)
    date = date_pattern.group(0)

    print('Saving as Excel at', datetime.now())
    print('Done!')
    input('Please select save location')
    directory = filedialog.askdirectory()
    with pd.ExcelWriter(directory + '/' + date + ' TML Report.xlsx') as writer:
            TML_report.to_excel(writer, sheet_name='TML', index=False)

    print("Done!")
    print("Saved at " + directory)

if __name__ == "__main__":
    main()
