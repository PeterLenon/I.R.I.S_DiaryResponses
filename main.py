import pandas as pd


def load_and_write_data(read_filepath, write_filepath):
    df = pd.read_excel(read_filepath)
    # print(df.head())
    to_keep = ['StartDate', 'RecipientLastName', 'RecipientFirstName', 'Q2', 'Q3_1', 'Q16', 'Q4', 'Q7', 'Q8', 'Q9',
               'Q10', 'Q11', 'Q12', 'Q13']
    df = df[to_keep]

    to_drop = ['Peter', 'Claire', 'Weslie', 'Selma', 'Manasi', 'Long-Jing']
    for name in to_drop:
        df = df[df.RecipientFirstName != name]
    df = df[df.RecipientFirstName.notnull()]
    df.to_excel(write_filepath)


analytics_read_file = "C:\\Users\gosho\OneDrive\Desktop\Expert+Panel-+Life+with+IRIS_July+6,+2023_20.17\Expert Panel- Life with IRIS_July 6, 2023_20.17.xlsx"
analystics_write_file = "C:\\Users\gosho\OneDrive\Desktop\Expert+Panel-+Life+with+IRIS_July+6,+2023_20.17\Expert Panel- processed.xlsx"
load_and_write_data(analytics_read_file, analystics_write_file)


