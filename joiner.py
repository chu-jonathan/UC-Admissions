import pandas as pd
import plotly as py
import plotly.express as px
import statsmodels.api as sm
import statsmodels.formula.api as smf
import sklearn as sk
import os

def ffill(df, col):
    df[col] = df[col].fillna(method='ffill')
    return df

def NaN_to_zero(df, col):
    df[col] = df[col].fillna(0)
    return df

def main():
    #gpas
    fill_columns = ["School", "City", "County", "Fall term"]
    dfs = []
    esuhsd_gpa_path = 'esuhsd_gpas'
    for filename in os.listdir(esuhsd_gpa_path):
        file_path = os.path.join(esuhsd_gpa_path, filename)
        if file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path, sheet_name="Sheet 1", engine='openpyxl')
            for col in fill_columns:
                df = ffill(df, col)
            df = df[df['City'] == 'San Jose']
            dfs.append(df)
    df = pd.concat(dfs)

    #admissions counts
    ad_count_fill_columns = ["School", "City", "Campus", "Fall Term"]
    df2s = []
    ad_count_path = 'esuhsd_ad_counts'
    for filename in os.listdir(ad_count_path):
        file_path = os.path.join(ad_count_path, filename)
        if file_path.endswith(".xlsx"):
            df2 = pd.read_excel(file_path, sheet_name="Sheet 1", engine='openpyxl')

            for col in ad_count_fill_columns:
                df2 = ffill(df2, col)
            df2 = df2[df2['City'] == 'San Jose']
            df2_pivot = df2.pivot_table(
                values='All', 
                index=['School', 'City','Campus', 'Fall Term'], 
                columns='Count', 
                aggfunc='first'  # Takes the first occurrence in case of duplicates
            )
            df2_pivot.reset_index(inplace=True)
            df2s.append(df2_pivot)
    df2_concat = pd.concat(df2s)
    print(df2_concat.columns)
    #addressing years no one got in
    add_NaNs = ["Adm", "App", "Enr"]
    for i in add_NaNs:
        df2_concat = NaN_to_zero(df2_concat, i)
    df2_concat['Adm_rate'] = df2_concat['Adm'] / df2_concat['App']

    gpa_Nans = ['Adm GPA', "Enrl GPA"]
    for i in gpa_Nans:
        df = NaN_to_zero(df, i)
    


    #export to excel to check
    df.rename(columns={'Fall term':'Fall Term'}, inplace=True)
    df.drop(columns=['Calculation1'], inplace=True)
    df.to_excel('outputs/gpa_concat.xlsx')
    df2_concat.to_excel('outputs/admission_counts_concat.xlsx')

    #join
    print(len(df))
    print(len(df2_concat))

    merged = pd.merge(df2_concat, df, how='left', on=['School', 'City', 'Campus','Fall Term'])   
    print(len(merged))

    #export to excel to check
    merged.to_excel('outputs/merged.xlsx')

    #universitywide model
    uv = merged[merged['Campus'] == 'Universitywide']
    uv.to_excel('outputs/uv.xlsx')


    #plot adm rate over time
    fig = px.line(data_frame=uv, x='Fall Term', y='Adm_rate', color='School')
    fig.write_html('plots/adm_rate_over_time.html')

    #plot adm over time
    fig = px.line(data_frame=uv, x='Fall Term', y='Adm', color='School')
    fig.write_html('plots/adm_over_time.html')

    #plot gpa over time
    fig = px.line(data_frame=uv, x='Fall Term', y='Adm GPA', color='School')
    fig.write_html('plots/gpa_over_time.html')

    #base model
    model = smf.mixedlm("Adm_rate ~ School + Q('App GPA')", data = uv, groups = uv['Fall Term'])
    result = model.fit()
    print(result.summary())

    
    model = smf.mixedlm("Adm ~ School + Q('App GPA') + App", data = uv, groups = uv['Fall Term'])
    result = model.fit()
    print(result.summary())

    #appbase model
    model = smf.mixedlm("App ~ School + Q('App GPA')", data = uv, groups = uv['Fall Term'])
    result = model.fit()
    print(result.summary())

    
if __name__ == "__main__":
    main()