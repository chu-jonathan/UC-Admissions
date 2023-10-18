import pandas as pd
import plotly as py
import statsmodels.api as sm
import sklearn as sk
import os

def ffill(df, col):
    df[col] = df[col].fillna(method='ffill')
    return df

def main():
    fill_columns = ["School", "City", "County", "Fall term"]
    dfs = []
    esuhsd_gpa_path = 'esuhsd_gpas'
    for filename in os.listdir(esuhsd_gpa_path):
        file_path = os.path.join(esuhsd_gpa_path, filename)
        if file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path, sheet_name="Sheet 1", engine='openpyxl')
            for col in fill_columns:
                df = ffill(df, col)
            dfs.append(df)
    df = pd.concat(dfs)
    print(df.tail(50))
    print(len(df))
    print(df.shape)
if __name__ == "__main__":
    main()