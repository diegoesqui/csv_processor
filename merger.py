#!/usr/bin/python3
import os
import pandas as pd
from datetime import datetime


def main():
    input_path = "./input"
    output_path = "./output"
    replacements = "replacement.xlsx"

    csv_files = [_ for _ in os.listdir(input_path) if _.endswith(".csv")]

    df = concat_files(csv_files, input_path)
    df = clean(df)
    df = set_replacements(df, replacements)
    generate_output(df, output_path, add_timestamp=True)

    print('Done!')
    return


def clean(df):
    df[['amount', 'accountBalance']] = df[['amount', 'accountBalance']].astype('float')
    df['dateVal'] = pd.to_datetime(df['dateVal'], format="%d/%m/%Y")
    df.drop(columns=['comment', 'pointer', 'dateOp'], inplace=True)
    return df


def concat_files(csv_files, input_path):
    df_list = []
    for file in csv_files:
        # read the csv file and store the data frame in the list
        df = pd.read_csv(os.path.join(input_path, file), sep=';', decimal=',', thousands=' ', dayfirst=True)
        df_list.append(df)

    # concatenate all the data frames into a single data frame
    df = pd.concat(df_list, ignore_index=True)

    # drop duplicates
    df = df.drop_duplicates()
    return df


def set_replacements(df, replacements):
    rep = pd.read_excel('replacements.xlsx')
    
    # If replacements file is empty
    if len(rep) == 0:
        # exit the function
        return df
    
    for row in rep.itertuples():
        if row.condition == 'is':
            df.loc[df[row.source_column] == row.source_value, row.destiny_column] = row.destiny_value
        elif row.condition == 'contains':
            df.loc[df[row.source_column].str.contains(row.source_value, case=False), row.destiny_column] = row.destiny_value

    df = df.drop_duplicates()
    return df


def generate_output(df, output_path, add_timestamp=False):
    # create the output folder if it doesn't exist
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    output_name = "data_merged"
    if add_timestamp:
        current_date = datetime.now().strftime("_%Y-%m-%d")
        output_name += current_date

    # export the data frame to a csv file in the output folder
    df.to_csv(os.path.join(output_path, output_name + ".csv"), index=False, sep=',')

    # export the data frame to a xlsx file in the output folder
    df.to_excel(os.path.join(output_path, output_name + ".xlsx"), index=False)
    return


if __name__ == "__main__":
    main()
    input("Press Enter to continue...")