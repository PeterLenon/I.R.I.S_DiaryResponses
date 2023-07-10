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
    df['Things_To_Do'] = ""
    df['Things_To_Not_Do'] = ""
    df['Essence_of_Those_Things'] = ""
    df['Low_IKIGAI_activity'] = ""
    df['High_IKIGAI_activity'] = ""
    df['IKIGAI_level'] = ""

    def things_To_or_Not(dataframe):
        for row in dataframe.index:
            dataframe['Things_To_Do'][row] = (str(dataframe['Q2'][row]) + "\n\n How IRIS helps: \n" + str(dataframe['Q10'][row]) + str(dataframe['Q11'][row])) if dataframe['Q9'][row] == 'Yes' else ""
            dataframe['Things_To_Not_Do'][row] = (str(dataframe['Q2'][row]) + "\n Why not: \n" + str(dataframe['Q13'][row])) if dataframe['Q9'][row] == 'No' else ""
            dataframe['Essence_of_Those_Things'][row] = dataframe['Q12'][row] if dataframe['Q9'][row] == 'Yes' else ""

    def graded_IKIGAI_activities(dataframe):
        for row in dataframe.index:
            if row != 0:
                dataframe['Low_IKIGAI_activity'][row] = dataframe['Q2'][row] if int(dataframe['Q3_1'][row]) < 50 else ""
                dataframe['High_IKIGAI_activity'][row] = dataframe['Q2'][row] if int(dataframe['Q3_1'][row]) >= 50 else ""
                dataframe['IKIGAI_level'][row] = str(dataframe['Q3_1'][row]) + '\n' + str(dataframe['Q16'][row])

    things_To_or_Not(df)
    graded_IKIGAI_activities(df)
    df.to_excel(write_filepath)



analytics_read_file = "C:\\Users\gosho\OneDrive\Desktop\Expert+Panel-+Life+with+IRIS_July+6,+2023_20.17\Expert Panel- Life with IRIS_July 6, 2023_20.17.xlsx"
analystics_write_file = "C:\\Users\gosho\OneDrive\Desktop\Expert+Panel-+Life+with+IRIS_July+6,+2023_20.17\Expert Panel- processed.xlsx"
load_and_write_data(analytics_read_file, analystics_write_file)


