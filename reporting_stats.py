import pandas as pd 
import numpy as np 
import os
import xlsxwriter
from Retrieve_emails import *
import shutil
from datetime import datetime
import statistics
import matplotlib
import matplotlib.pyplot as plt
from io import BytesIO
import sys
# %matplotlib inline
# matplotlib.style.use('ggplot')

year = str(input("Enter year:"))
# year = '2020'
path = os.path.abspath(os.path.dirname(__file__))
try:
    shutil.rmtree(path + '\\data')
except:
        pass
os.mkdir(path + '\\data\\')
retrieve_emails(path, year)

df = pd.read_csv(path + "\\data\\" + "outlook_emails_" + year + ".csv")
df = df[~df["Full path"].str.contains('Spam')].reset_index(drop= True)
df = df[~df["Full path"].str.contains('spam')].reset_index(drop= True)

def build_new_id(x):
    if x['ID'] is np.nan or x['ID'] == 'nan' or x['ID'] == 'NaN':
        return(x['conversationID'])
    else: return(x['ID'])

df['ID'] = df['ID'].apply(str)
df['new_ID'] = df.apply(lambda row: build_new_id(row), axis = 1)

 # Number of queries
num_distinct_queries = df['ID'].drop_duplicates().count()
# + df[df.ID.isna()]['conversationID'].drop_duplicates().count()

 # Number of media
num_media = df[df.is_media == True]['new_ID'].drop_duplicates().count()

 # Number of final replies
num_final_replies = len(df[df.is_final_reply == True]) + num_media

 # Number of direct replies
num_direct_replies = len(df.loc[(df.is_final_reply == True) & (df.is_direct_reply == True)])

 # Number of queries from CRM
num_crm = len(df.loc[(df.is_final_reply == True) & (df.is_crm == True)]['ID'].drop_duplicates())

# Number of follow-ups
followups = df[df.is_final_reply == True].groupby(['Recipients', 'ID']).size().reset_index(name='counts')
num_followups = (followups[followups.counts > 1]['counts']-1).sum()

df['time_format'] = df.apply(lambda row: datetime.strptime(row['Received Time'], '%Y/%m/%d %H:%M:%S'), axis=1)

    # Compute average time of answer
new_df = df[~df["Full path"].str.contains('Additional')]
new_df = new_df[~new_df["Full path"].str.contains('Media')]
new_df = new_df.groupby(['new_ID'])['time_format'].agg([max, min])

def weekend_excl(x):
    delta_diff = x['max'] - x['min']
    delta_we = np.busday_count(x['min'].date(), x['max'].date())
    if int(delta_diff.days) > delta_we:
        delta_diff = delta_diff - pd.Timedelta(str(delta_diff.days) + ' days') \
            + pd.Timedelta(str(delta_we) + ' days') 
    return(delta_diff)

new_df['delta'] = new_df.apply(lambda x: weekend_excl(x), axis = 1)

answer_time = new_df['delta']
answer_time = answer_time[(answer_time < pd.Timedelta("30 days"))]
answer_time = str(np.mean(answer_time[(answer_time > pd.Timedelta("0 days"))]))[:9] + ' hours'

    # Compute average time to answer media
new_df = df[df.is_media == True].groupby(['new_ID'])['time_format'].agg([max, min])
new_df['delta'] = new_df.apply(lambda x: weekend_excl(x), axis = 1)
answer_time_media = new_df['delta']
all_media = answer_time_media.apply(lambda x: str(x))

answer_time_media = answer_time_media[(answer_time_media < pd.Timedelta("15 days"))]

media_stat = answer_time_media[(answer_time_media > pd.Timedelta("0 days"))].describe()
media_stat = media_stat.apply(lambda x: str(x))[1:]
plot_media = (answer_time_media[(answer_time_media > pd.Timedelta("0 days"))] / \
    pd.Timedelta(hours=1))
    
plot_media = plt.hist(x=plot_media, bins=range(0, 60, 2), color='#0504aa',
                            alpha=0.7, rwidth=0.85)
                            
plot_media = pd.DataFrame({'Hours': plot_media[1][:-1], \
        'Number of queries': plot_media[0]}, \
        columns = list(['Hours', 'Number of queries']))
plot_media = plot_media.set_index(plot_media.columns[0])

# answer_time_media = str(
#     statistics.median(answer_time_media[(answer_time_media > pd.Timedelta("0 days"))])
#     )[:9] + ' hours'

# answer_time_media = str(
#     np.mean(answer_time_media[(answer_time_media > pd.Timedelta("0 days"))])
#     )[:9] + ' hours'
# answer_time_media.to_csv('checkmedia.csv')

d = {'N. of emails (internal exchanges included)': len(df),
     'N. of distinct queries (queries without ID not counted)': num_distinct_queries,
     'N. of queries from CRM': num_crm,
     'N. of follow-ups': num_followups,
     'N. of final replies (follow-ups included)': num_final_replies, 
     'N. of direct replies to user (without help of BA)': num_direct_replies,
     'Relative n. of direct replies': str(round(100*num_direct_replies/num_final_replies,2)) + ' %',
     'N. of media queries': num_media,
     'Average time needed to close a query': answer_time}

general_stat = pd.DataFrame(d, index = [year]).transpose()

 # Add BA 
BAs = pd.DataFrame(list(df['Full path'].apply(lambda x: x.split("\\")[-1:])),\
     columns = list(['BA']))
df = pd.concat([df, BAs], axis = 1)

 # Stat on BAs
final_replies_BA = df[df.is_final_reply == True].groupby(['BA']).size()

final_replies_BA_help = df[df.is_media == False][df.is_final_reply == True][df.is_direct_reply == False]\
    .groupby(['BA']).size()
final_replies_BA_direct = df[df.is_final_reply == True][df.is_direct_reply == True]\
    .groupby(['BA']).size()


    # Compute average time of answer
new_df = df[df.is_direct_reply == False].groupby(['BA','new_ID'])['time_format'].agg([max, min])
new_df[1] = new_df.apply(lambda x: weekend_excl(x), axis = 1)
answer_time = new_df[1]
answer_time = answer_time[(answer_time < pd.Timedelta("40 days"))]
answer_time = answer_time[(answer_time > pd.Timedelta("0 days"))]
answer_time = answer_time.groupby(level=0).agg(np.mean)

final_replies_BA_help = pd.concat([final_replies_BA_help, answer_time], axis=1, join="inner")
final_replies_BA_help[1] = final_replies_BA_help[1].apply(lambda x: str(x)[:9] + ' hours')
final_replies_BA_help = final_replies_BA_help.rename(columns={0: "Number of queries", 1: "Avg time needed"})

final_replies_BA_direct = pd.DataFrame(final_replies_BA_direct, columns = list(['Direct replies']))


    # Number of queries over time per BA
count_BA = df[df.is_final_reply == True][['BA', 'months']]\
            .groupby(['BA', 'months']).size()\
            .reset_index()
count_BA.columns = list(['BA', 'month', 'counts'])

count_BA = count_BA.pivot_table(index=['BA'],columns='month', values='counts')


    # Direct replies over time
DR_over_time = pd.DataFrame(df[df.is_direct_reply == True][df.is_final_reply == True]\
            .groupby(['months']).size())
DR_over_time.columns = list(['N. replies'])

# list of dataframes and sheet names
dfs = [general_stat, 
       plot_media, all_media, 
       final_replies_BA_help,
       final_replies_BA_direct,
       count_BA,
       DR_over_time]
sheets = ['General Statistics', 'Media Statistics',
    'All media queries', 'Queries per BA', \
    'Queries answered directly', 'Queries BA per month',
    'Direct replies over time']    

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# run function to write in differrent sheets in Excel
def dfs_tabs(df_list, sheet_list, file_name):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0)   
   
    writer.save()

dfs_tabs(dfs, sheets, desktop + '\\report_statistics_' + year + '.xlsx')


print("Annual report " + year + ' built. You can find it in your desktop!')
input('Click Enter to close')
shutil.rmtree(path + '\\data')
