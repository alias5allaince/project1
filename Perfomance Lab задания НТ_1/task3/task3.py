
import pandas as pd
from datetime import datetime

check_date1="28.04.2019"
check_date2="29.04.2020"


dateFormatter = "%d.%m.%Y"
valid_date1=datetime.strptime(check_date1, dateFormatter)
valid_date2=datetime.strptime(check_date2, dateFormatter)

df_log = pd.read_excel('Log.xlsx',sheet_name = 'Log')

df1=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2)].values# вывод массива в диапазоне дат (начало и конец лога+)
df2=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['action'] == 'top up')].values # вывод массива попыток налить воду в периоде
df3=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['status'] == 'фейл')].values # массив с количеством ошибок
df4=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['status'] == 'успех') & (df_log['action'] == 'top up')].values # массив с успешным долитием воды в бочку
df5=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['status'] == 'фейл') & (df_log['action'] == 'top up')].values # массив с НЕ налитием воды в бочку
df6=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['status'] == 'успех') & (df_log['action'] == 'scoop')].values # массив с успешным забором воды из бочки
df7=df_log.loc[(df_log['date_value'] > valid_date1) & (df_log['date_value'] < valid_date2) & (df_log['status'] == 'фейл') & (df_log['action'] == 'scoop')].values # массив с НЕуспешным забором воды из бочки


if len(df1)>0: #если больше 0, то можно искать, иначе ошибка, нет действий в интервале.
    print("Количество попыток налить воду в бочку составляет" + " " + str(len(df2)) + " раз")
    fail=round(len(df3)/len(df1),3)*100
    print("Процент допущенных ошибок за период составляет" + " " + str(fail) +"%")
    i=0
    a=0
    for r in df4:
        a+=r[4]
        i=i+1
    print("Объем налитой в бочку воды составляет" + " " + str(a) + " " + "литров")
    b=0
    for r in df5:
        b+=r[4]
        i=i+1
    print("Объем НЕ налитой в бочку воды составляет" + " " + str(b) + " " + "литров")
    c=0
    for r in df6:
        c+=r[4]
        i=i+1
    print("Объем зачерпнутой воды составляет" + " " + str(c) + " " + "литров")
    d=0
    for r in df7:
        d+=r[4]
        i=i+1
    print("Объем неуспешного забора воды составляет" + " " + str(d) + " " + "литров")
    print("Объём воды в бочке на начало периода составляет" + " " + str(df1[0,0]) + " " + "литров")
    print("Объём воды в бочке на конец периода составляет" + " " + str(df1[len(df1)-1,6]) + " " + "литров")
else:
    print("ЗА выбранный интервал событий не происходило")
