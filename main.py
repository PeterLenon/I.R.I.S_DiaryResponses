import matplotlib.dates
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import plotly.graph_objects as go

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

    df['IKIGAI level to use IRIS'] = ""
    df['IKIGAI level to not use IRIS'] = ""
    df['Appropriate time to use IRIS'] = ""
    df['Inappropriate time to use IRIS'] = ""
    df['Activity and Occasion'] = ""



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

    def IKIGAI_to_use_IRIS(dataframe):
        for row in dataframe.index:
            dataframe['IKIGAI level to use IRIS'][row] = dataframe['Q3_1'][row] if dataframe['Q9'][row] == 'Yes' else ""
            dataframe['IKIGAI level to not use IRIS'][row] = dataframe['Q3_1'][row] if dataframe['Q9'][row] =='No' else ""

    def appropriate_time_to_use_IRIS(dataframe):
        for row in dataframe.index:
            if row != 0:
                dataframe['Appropriate time to use IRIS'][row] = dataframe['StartDate'][row].time() if dataframe['Q9'][row] == 'Yes' else ""
                dataframe['Inappropriate time to use IRIS'][row] = dataframe['StartDate'][row].time() if dataframe['Q9'][row] == 'No' else ""

    def activity_and_occasion(dataframe):
        for row in dataframe.index:
            dataframe['Activity and Occasion'][row] = '' if dataframe['Q9'][row] != 'Yes' else ('Activity: ' + (dataframe['Q11'][row] if dataframe['Q10'][row] == 'Others, please specify:' else dataframe['Q10'][row]) + '\n' + 'Occasion: ' + dataframe['Q2'][row])

    def time_to_ikigai_stats(dataframe):
        ap_x_values = []
        ap_y_values = []
        in_x_values = []
        in_y_values = []
        for row in dataframe.index:
            if row != 0:
                if dataframe['Appropriate time to use IRIS'][row] != "":
                    ap_x_values.append(datetime.combine(datetime.today(), dataframe['Appropriate time to use IRIS'][row]))
                    ap_y_values.append(dataframe['IKIGAI level to use IRIS'][row])

                if dataframe['Inappropriate time to use IRIS'][row] != "":
                    in_x_values.append(datetime.combine(datetime.today(), dataframe['Inappropriate time to use IRIS'][row]))
                    in_y_values.append(dataframe['IKIGAI level to not use IRIS'][row])

        plt.scatter(ap_x_values, ap_y_values)
        plt.scatter(in_x_values, in_y_values)
        plt.grid()
        plt.gcf().autofmt_xdate()
        plt.show()


    things_To_or_Not(df)
    graded_IKIGAI_activities(df)
    IKIGAI_to_use_IRIS(df)
    appropriate_time_to_use_IRIS(df)
    activity_and_occasion(df)
    time_to_ikigai_stats(df)

    df.to_excel(write_filepath)


analytics_read_file = "C:\\Users\gosho\OneDrive\Desktop\R-HouseFiles\Expert+Panel-+Life+with+IRIS_July+10,+2023_11.13\Expert Panel- Life with IRIS_July 10, 2023_11.13.xlsx"
analystics_write_file = "C:\\Users\gosho\OneDrive\Desktop\R-HouseFiles\Expert+Panel-+Life+with+IRIS_July+10,+2023_11.13\Expert Panel- processed_July 10.xlsx"
load_and_write_data(analytics_read_file, analystics_write_file)


