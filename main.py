import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import pyplot as plt, dates
from datetime import datetime
import numpy as np
import random
import seaborn as sns
import matplotlib.dates as md

def load_and_write_data(read_filepath, write_filepath):
    df = pd.read_excel(read_filepath)
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

    def time_to_ikigai_plots(dataframe):
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

        fig, ax = plt.subplots()
        ax.scatter(ap_x_values, ap_y_values, alpha=0.5)
        ax.scatter(in_x_values, in_y_values, alpha=0.5)
        plt.grid()
        plt.gcf().autofmt_xdate()
        ax.xaxis.set_major_formatter(dates.DateFormatter('%H:%M'))
        plt.ylabel('Effect on IKIGAI %')
        plt.title('Appropriate(blue) /Inappropriate(orange) times vs Ikigai Level')
        plt.savefig('Appropriate and Inappropriate times vs Ikigai Level.png')
        plt.show()


    def IRIS_activity_vs_time_scatterplot(dataframe):
        y_activity_labels = []
        x_time_labels = []

        for row in dataframe.index:
            if row != 0 and dataframe['Q9'][row] == 'Yes':
                y_activity_labels.append(dataframe['Q10'][row] if dataframe['Q10'][row] in ['Reflection activity', 'General chat about ikigai', 'Photograph activity'] else 'Other')
                x_time_labels.append(datetime.combine(datetime.today(), dataframe['StartDate'][row].time()))

        fig, ax = plt.subplots()
        ax.scatter(x_time_labels, y_activity_labels, alpha=0.5)
        plt.tick_params(rotation=45)
        plt.gcf().autofmt_xdate()
        ax.xaxis.set_major_formatter(dates.DateFormatter('%H:%M'))
        plt.grid()
        plt.savefig('IRIS_activity_vs_time.png')
        plt.show()

    def IRIS_activity_vs_ikigai_level(dataframe):
        y_activity_labels = []
        x_corresponding_ikigai_level = []

        for row in dataframe.index:
            if dataframe['Q9'][row] == 'Yes':
                Q10list = dataframe['Q10'][row].split(',')
                for activity in Q10list:
                    if activity != ' please specify:':
                        y_activity_labels.append(activity)
                        x_corresponding_ikigai_level.append(dataframe['IKIGAI level to use IRIS'][row])

        plt.scatter(x_corresponding_ikigai_level, y_activity_labels, alpha=0.5)
        plt.xlabel('How much current activity affects IKIGAI %')
        plt.grid()
        plt.savefig('IRIS_activity_vs_ikigai_level.png')
        plt.title('Appropriate IRIS activity vs corresponding IKIGAI Level')
        plt.show()

    def most_and_least_utilised_activity(dataframe):
        activity_freq_dict = {'Reflection activity': 0, 'General chat about ikigai': 0, 'Photograph activity':0 , 'Other': 0}
        for row in dataframe.index:
            if dataframe['Q10'][row] == 'Reflection activity':
                activity_freq_dict['Reflection activity'] = activity_freq_dict['Reflection activity']+1
            elif dataframe['Q10'][row] == 'General chat about ikigai':
                activity_freq_dict['General chat about ikigai'] = activity_freq_dict['General chat about ikigai'] + 1
            elif dataframe['Q10'][row] == 'Photograph activity':
                activity_freq_dict['Photograph activity'] = activity_freq_dict['Photograph activity'] + 1
            elif dataframe['Q10'][row] == 'Others, please specify:':
                activity_freq_dict['Other'] = activity_freq_dict['Other'] + 1

        y = []
        activity_labels = []
        for activity_label in activity_freq_dict:
            activity_labels.append(activity_label)
            y.append(activity_freq_dict[activity_label])

        plt.pie(y, labels = activity_labels, autopct= '%1.1f%%')
        plt.title('% Most and Least suggested activity')
        plt.savefig('most_and_least_utilised_activity.png')
        plt.show()

    def who_are_you_with_vs_IRIS_usecounts_plots(dataframe):
        yes_person_freq_dict = {'Alone': 0,
                            'Spouse': 0,
                            'Children': 0,
                            'Grandchildren': 0,
                            'Friends': 0,
                            'Nurse/Care Partners': 0,
                            'Others':0
                            }

        no_person_freq_dict = {'Alone': 0,
                                'Spouse': 0,
                                'Children': 0,
                                'Grandchildren': 0,
                                'Friends': 0,
                                'Nurse/Care Partners': 0,
                                'Others': 0
                                }

        for row in dataframe.index:
            Q7list = dataframe['Q7'][row].split(',')
            if dataframe['Q9'][row] == 'Yes':
                for key in yes_person_freq_dict:
                    if key in Q7list:
                        yes_person_freq_dict[key] = yes_person_freq_dict[key] + 1
            elif dataframe['Q9'][row] =='No':
                for key in no_person_freq_dict:
                    if key in Q7list:
                        no_person_freq_dict[key] = no_person_freq_dict[key]+1



        plt.bar(yes_person_freq_dict.keys(), yes_person_freq_dict.values(), color='blue', width=0.3)
        plt.title('IRIS use count vs who you are with')
        plt.tick_params(rotation=45)
        plt.ylabel('Number of times')
        plt.savefig('Iris use yes and no graph.png')
        plt.savefig('Persons and IRIS_use.png')
        plt.show()

        plt.bar(no_person_freq_dict.keys(), no_person_freq_dict.values(), color='orange', width=0.3)
        plt.title('no IRIS use vs who you are with')
        plt.tick_params(rotation=45)
        plt.ylabel('Number of times')
        plt.savefig('Persons and NO IRIS_use.png')
        plt.show()

    def IRIS_activity_vs_who_are_you_with(dataframe):
        person_to_index_dict = {'Alone': 0,
                                'Spouse': 1,
                                'Children': 2,
                                'Grandchildren': 3,
                                'Friends': 4,
                                'Nurse/Care Partners': 5,
                                'Others': 6
                                }

        IRIS_activity_to_person_dict = {'Others': [0, 0, 0, 0, 0, 0, 0],
                                        'Reflection activity': [0, 0, 0, 0, 0, 0, 0],
                                        'General chat about ikigai': [0, 0, 0, 0, 0, 0, 0],
                                        'Photograph activity': [0, 0, 0, 0, 0, 0, 0]
                                        }
        for row in dataframe.index:
            if dataframe['Q9'][row] == 'Yes':
                IRIS_activity_list = dataframe['Q10'][row].split(',')
                # IRIS_activity_list.pop('please specify:')
                person_list = dataframe['Q7'][row].split(',')

                for activity in IRIS_activity_list:
                    if activity != ' please specify:':
                        LIST = IRIS_activity_to_person_dict[activity]
                        for person in person_list:
                            if person != ' please specify:':
                                index = person_to_index_dict[person]
                                LIST[index] = LIST[index] + 1


        for person in person_to_index_dict.keys():
            for activity in IRIS_activity_to_person_dict.keys():
                text = IRIS_activity_to_person_dict[activity][person_to_index_dict[person]]
                if text != 0:
                    plt.scatter(person, activity, s=int(text)*100, c = 40, marker='o', alpha=0.5)
                    plt.annotate(text, (person, activity))

        plt.grid()
        plt.title('who are you with vs desired IRIS activity')
        plt.savefig('who are you with vs IRIS activity.png')
        plt.show()


    things_To_or_Not(df)
    graded_IKIGAI_activities(df)
    IKIGAI_to_use_IRIS(df)
    appropriate_time_to_use_IRIS(df)
    activity_and_occasion(df)
    time_to_ikigai_plots(df)
    IRIS_activity_vs_time_scatterplot(df)
    IRIS_activity_vs_ikigai_level(df)
    most_and_least_utilised_activity(df)
    who_are_you_with_vs_IRIS_usecounts_plots(df)
    IRIS_activity_vs_who_are_you_with(df)

    df.to_excel(write_filepath)

analytics_read_file = "C:\\Users\gosho\OneDrive\Desktop\R-HouseFiles\Expert+Panel-+Life+with+IRIS_July+17,+2023_09.50\Expert Panel- Life with IRIS_July 17, 2023_09.50.xlsx"
analytics_write_file = "C:\\Users\gosho\OneDrive\Desktop\R-HouseFiles\Expert+Panel-+Life+with+IRIS_July+17,+2023_09.50\Expert Panel- Life with IRIS_July 17-processed.xlsx"
load_and_write_data(analytics_read_file, analytics_write_file)


