from tkinter import *
import pandas as pd
import xlsxwriter
import tkinter.ttk as ttk
from datetime import datetime
from tkinter import messagebox as mb
import time
from random import randint
import os

def creating():# идет проверка ввода даты, формата даты, наличие ввода в остальных полях, далее в зависимости от результатов происходит добавление заявки, или вывод информационного окна.
    try:
        date_request=datetime.now()
        dateFormatter = "%d.%m.%Y"
        date_value = datetime.strptime(date.get(), dateFormatter)
        fio_manager_value = fio_manager.get()
        Delivery_method_value = Delivery_method.get()
        Payment_option_value = Payment_option.get()
        fio_klient_value = fio_klient.get()
        phone_number_value = phone_number.get()
        e_mail_value = e_mail.get()
        adress_value = adress.get()
        hard_long_request_value=hard_long_request.get()# значение 0 если заявка в пределах МКАД, значение 4 если надо ехать за МКАД. типо 4 часа, заменяет 4 заявки.
        FOP_number_application_value = FOP_number_application.get()
        comments_value = comments.get(1.0, END)
        if date_value>=date_request and len(fio_manager_value)!=0 and len(Delivery_method_value)!=1 and len(Payment_option_value)!=1 and len(fio_klient_value)!=0 and len(phone_number_value)!=0 and len(adress_value)!=0 and len(FOP_number_application_value)!=0:
            df_delivery = pd.read_excel('Base\Delivery_base.xlsx', sheet_name = 'Sheet1')
            quantity_delivery_on_date=len(df_delivery.loc[df_delivery['date_value'] == date_value])#получили кол-во заявок на выбранный день
            df_max_quantity_delivery = pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
            max_quantity_delivery=int(df_max_quantity_delivery.iloc[0][0])#из ексель файла берётся единственное значение - максимальное количество заявок в день, устанавливается в зависимости от кол-ва выездников
            if quantity_delivery_on_date<max_quantity_delivery:
                if hard_long_request_value==4 and max_quantity_delivery - quantity_delivery_on_date>=3:
                    i=0
                    while i<hard_long_request_value:# если заявка требует выезда за МКАД, она дублируется 4 раза.
                        df_delivery = pd.read_excel('Base\Delivery_base.xlsx', sheet_name = 'Sheet1')
                        df_delivery.loc[len(df_delivery)] = [date_request, fio_manager_value, date_value, Delivery_method_value,Payment_option_value,fio_klient_value,e_mail_value,phone_number_value,adress_value,FOP_number_application_value,comments_value]
                        df_delivery.to_excel("Base\Delivery_base.xlsx",sheet_name='Sheet1',index=False)
                        i=i+1
                    mb.showinfo("Урааааа!!!",'Заявка успешно добавлена =)')
                    date.delete(0, 'end')# очищаем все поля после добавления заявки
                    fio_manager.delete(0, 'end')
                    date.delete(0, 'end')
                    fio_klient.delete(0, 'end')
                    e_mail.delete(0, 'end')
                    Delivery_method.set(0)
                    Payment_option.set(0)
                    hard_long_request.set(0)
                    phone_number.delete(0, 'end')
                    adress.delete(0, 'end')
                    FOP_number_application.delete(0, 'end')
                    comments.delete(1.0, 'end')
                elif hard_long_request_value==4 and max_quantity_delivery - quantity_delivery_on_date<=3:
                    mb.showerror("Жаль...",'Требуется выезд за МКАД, недостаточно времени')#вот тут надо подумать, если нет мест вообще выводится это, а должно быть другое сообщение
                elif hard_long_request_value!=4 and quantity_delivery_on_date<max_quantity_delivery:#сравнили количество текущих заявок на выбранный день с максимальным, и если меньше - записали заявку
                    df_delivery.loc[len(df_delivery)] = [date_request, fio_manager_value, date_value, Delivery_method_value,Payment_option_value,fio_klient_value,e_mail_value,phone_number_value,adress_value,FOP_number_application_value,comments_value]
                    df_delivery.to_excel("Base\Delivery_base.xlsx",sheet_name='Sheet1',index=False)
                    mb.showinfo("Урааааа!!!",'Заявка успешно добавлена =)')
                    date.delete(0, 'end')# очищаем все поля после добавления заявки
                    fio_manager.delete(0, 'end')
                    date.delete(0, 'end')
                    fio_klient.delete(0, 'end')
                    e_mail.delete(0, 'end')
                    Delivery_method.set(0)
                    Payment_option.set(0)
                    hard_long_request.set(0)
                    phone_number.delete(0, 'end')
                    adress.delete(0, 'end')
                    FOP_number_application.delete(0, 'end')
                    comments.delete(1.0, 'end')
            else:
                mb.showerror("Жаль...",'Сорри, добавить заявку нельзя - мест нет =(')
        else:
            if date_value<=date_request:
                mb.showerror("Ошибка", "Необходимо ввести будущую дату, не прошлую. День в день заявки также не принимаются")
            else:
                mb.showerror("Ошибка записи", "Необходимо заполнить обязательные поля, без их заполнения заявка не создастся")
    except ValueError:
        mb.showwarning("Внимание", "Дата должна быть введена в формате ДД.ММ.ГГГГ")
    except xlsxwriter.exceptions.FileCreateError:
        mb.showwarning("Внимание", "Ошибка записи, файл временно недоступен, просьба повторить через 5-10 минут, спасибо.")
        
def give_number(): #получаем введеные в форме ФИО и продукт, в зависимости от продукта идет ветвление, обращение к екселю, "вытаскивание" последнего значения, увеличение его на 1, и запись обратно, одновременно - выод инфо в поле.
    FIO=fio_get_number.get()
    Product=choose_produkt.get()
    if Product=='Спроси Врача (Спроси Врача Лайт)' and len(FIO)!=0:# нет ФИО или не выбран продукт - вывод об ошибке.
        df_give_number = pd.read_excel('Numbers\Sprosi_vracha.xlsx',sheet_name = 'Sprosi_vracha')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])#из файла берется значение текущего года - для формирования номера
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        df_give_number.to_excel('Numbers\Sprosi_vracha.xlsx',sheet_name = 'Sprosi_vracha',index=False)
        str_number=str(number+1)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) #в это поле вывод полученного номера, для каждого варианта свой вариант вывода.
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер Спроси Врача\n18' + year + '-51 КМПСХ '+ str_number[1:])   
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=='Нет болезням' and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\_Net_boleznyam.xlsx',sheet_name = 'Net_boleznyam')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\_Net_boleznyam.xlsx',sheet_name = 'Net_boleznyam',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD)
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер Нет болезням\n18' + year + '-51 ПНБ '+ str_number[1:])
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Доктор Лайк (франшиза 0)" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер Доктор Лайк\n18' + year + '-51 ONDL 00'+ str_number[1:])
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Доктор Лайк (франшиза 30)" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер Доктор Лайк\n18' + year + '-51 ONDL 30'+ str_number[1:]) 
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Доктор Лайк (франшиза 50)" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\Doctor_like.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер Доктор Лайк\n18' + year + '-51 ONDL 50'+ str_number[1:])  
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="ИФЛ персональное решение" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\Personalka_PP.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\Personalka_PP.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для персонального решения\n18' + year + '-51 PP '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Онкопомощь" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\ONKO.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\ONKO.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для Онкопомощи\n18' + year + '-51 CC '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="ИФЛ простое/оптимальное решение" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\Korobki_PKS.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\Korobki_PKS.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для Простого/оптимального решения\n18' + year + '-51 PKS '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="ВПМЖ" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\VPMZH.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\VPMZH.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для полиса ВПМЖ\n18' + year + '-51 SL '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Автокаско" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\AVTOKASKO.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\AVTOKASKO.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для полиса АВТОКАСКО\n18' + year + '-51 MP '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)
    elif Product=="Несчастный случай" and len(FIO)!=0:
        df_give_number = pd.read_excel('Numbers\_NS.xlsx',sheet_name = 'Sheet1')
        number = df_give_number['№'].values[-1]
        df_year=pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
        year = str(df_year.iloc[4][0])
        df_give_number.loc[len(df_give_number)] = [number+1, FIO]
        str_number=str(number+1)
        df_give_number.to_excel('Numbers\_NS.xlsx',sheet_name = 'Sheet1',index=False)
        outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
        outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)
        outlbl_give_number.insert(1.0, 'Привет, '+ FIO +', твой номер для полиса НС\n18' + year + '-51 KA '+ str_number)
        make_menu(give_number_frame)
        outlbl_give_number.bind("<Button-3><ButtonRelease-3>", show_menu)     
    else:
        mb.showwarning(
        "Ошибка", 
        "Необходимо выбрать продукт и указать ФИО!")
  
def check_delivery():# проверяет правильность введенной даты, выгружает массив из екселя по соответствию поля даты выезда, выводит в отдельном окне все заявки на дату.
    try:
        valid_date = time.strptime(date.get(), '%d.%m.%Y')
        dateFormatter = "%d.%m.%Y"
        check_date = datetime.strptime(date.get(), dateFormatter)
        df_delivery_check = pd.read_excel('Base\Delivery_base.xlsx', sheet_name = 'Sheet1')# считали весь файл
        quantity_delivery_on_date=len(df_delivery_check.loc[df_delivery_check['date_value'] == check_date]) #количество строк по условию id=3
        df_all_request=df_delivery_check.loc[df_delivery_check['date_value'] == check_date].values #сами строки по условию соответствия выбранной дате, массив
        if quantity_delivery_on_date>0:
            show_all_request = Toplevel(window)
            show_all_request.title("Check request") 
            show_all_request.geometry('1200x700+100+50')
            def onFrameConfigure(canvas):#я хз как это работает, но без этой функции прокрутка не работает
                canvas.configure(scrollregion=canvas.bbox("all"))
            canvas = Canvas(show_all_request, borderwidth=1)
            frame =  Frame(canvas)
            vsb = Scrollbar(show_all_request, orient="vertical", command=canvas.yview)
            hsb = Scrollbar(show_all_request, orient="horizontal", command=canvas.xview)
            canvas.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
            vsb.pack(side="right", fill="y")
            hsb.pack(side='bottom',fill="x")
            canvas.pack(side="left", fill="both", expand=True)
            canvas.create_window((0,0), window=frame, anchor="nw")
            frame.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))#я хз как это работает, но без этой функции прокрутка не работает
            i=0
            while i<quantity_delivery_on_date:# при условии что заявки есть, запускается цикл вытаскивания массивов r (каждая отдельная заявка) из общего массива h (один массив из всех заявок по условию).
                for r in df_all_request: # первые 11 строк - заголовки, дальше добавляется количество строк исходя из значения g, которое привязано к i
                    Label(frame, width=15,height=3, bg='#CDE4F2',font='Constantia 9', text='Дата создания \n заявки').grid(row=0,column=0,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=20,height=3, bg='#CDE4F2',font='Constantia 9', text='ФИО \n менеджера').grid(row=0,column=1,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=15,height=3, bg='#CDE4F2',font='Constantia 9', text='Дата \n выезда к клиенту').grid(row=0,column=2,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=10,height=3, bg='#CDE4F2',font='Constantia 9', text='Способ доставки').grid(row=0,column=3,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=20,height=3, bg='#CDE4F2',font='Constantia 9', text='Вопрос оплаты').grid(row=0,column=4,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=20,height=3, bg='#CDE4F2',font='Constantia 9', text='ФИО клиента').grid(row=0,column=5,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=10,height=3, bg='#CDE4F2',font='Constantia 9', text='E-mail клиента').grid(row=0,column=6,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=10,height=3, bg='#CDE4F2',font='Constantia 9', text='Номер телефона \n клиента').grid(row=0,column=7,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=20,height=3, bg='#CDE4F2',font='Constantia 9', text='Адрес доставки \n полиса').grid(row=0,column=8,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=25,height=3, bg='#CDE4F2',font='Constantia 9', text='Номер заявки в фоп \n (где/у кого находится полис)').grid(row=0,column=9,sticky="nsew", padx=1, pady=1)
                    Label(frame, width=20,height=3, bg='#CDE4F2',font='Constantia 9', text='Комментарии').grid(row=0,column=10,sticky="nsew", padx=1, pady=1)

                    text1=Text(frame, width=20,height=1,bg='#CDE4F2',font='Georgiaa 9')#+1, чтобы первая строка заголовок, а вторая и последующие - данные
                    text1.grid(row=i+1,column=0,sticky="nsew", padx=1, pady=1)
                    text1.insert(1.0, r[0].strftime('%d.%m.%Y'))

                    text2=Text(frame,width=20,height=2,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text2.grid(row=i+1,column=1,sticky="nsew", padx=1, pady=1)
                    text2.insert(1.0, r[1])

                    text3=Text(frame,width=20,height=1, bg='#CDE4F2',font='Georgiaa 9')
                    text3.grid(row=i+1,column=2,sticky="nsew", padx=1, pady=1)
                    text3.insert(1.0, r[2].strftime('%d.%m.%Y'))

                    text4=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text4.grid(row=i+1,column=3,sticky="nsew", padx=1, pady=1)
                    text4.insert(1.0, r[3])

                    text5=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text5.grid(row=i+1,column=4,sticky="nsew", padx=1, pady=1)
                    text5.insert(1.0, r[4])
                    
                    text6=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text6.grid(row=i+1,column=5,sticky="nsew", padx=1, pady=1)
                    text6.insert(1.0, r[5])
                    
                    text7=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text7.grid(row=i+1,column=6,sticky="nsew", padx=1, pady=1)
                    text7.insert(1.0, r[6])
                    
                    text8=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text8.grid(row=i+1,column=7,sticky="nsew", padx=1, pady=1)
                    text8.insert(1.0, r[7])
                    
                    text9=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text9.grid(row=i+1,column=8,sticky="nsew", padx=1, pady=1)
                    text9.insert(1.0, r[8])
                    
                    text10=Text(frame,width=20,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text10.grid(row=i+1,column=9,sticky="nsew", padx=1, pady=1)
                    text10.insert(1.0, r[9])
                    
                    
                    text11=Text(frame,width=40,height=3,wrap=WORD, bg='#CDE4F2',font='Georgiaa 9')
                    text11.grid(row=i+1,column=10,sticky="nsew", padx=1, pady=1)
                    text11.insert(1.0, r[10])
                    
                    i=i+1
                    date.delete(0, 'end')
        else:
            mb.showinfo("Для инфо",'Заявок на выбранную дату нет.')
            date.delete(0, 'end')
    except ValueError:
        mb.showwarning('Ошибка', 'Дата должна быть введена в формате ДД.ММ.ГГГГ')
        date.delete(0, 'end')
        
def print_aforism():# просто выводим на печать рандомный афоризм из файла ексель - ради самообразования:)
    window.after(30000, print_aforism)# создается повтор каждые 30 секунд, т.к. ссылается сама на себя 30000
    df_aforism = pd.read_excel('Other\Aforism.xlsx', sheet_name = 'Sheet1')
    max_rand=len(df_aforism)
    R=randint(1,max_rand)
    Print_aforism=df_aforism.iloc[R][1] #выбор рандомной строчки от 1 до последней, последняя строчка определяется по длине списка
    show_aforism.delete(1.0, END)
    show_aforism.insert(1.0, Print_aforism)

def save_file(): # загружает адрес на диске из файла, создает там папку по ФИО+дата для сохранения документов по доставке
    try:
        date_value = date.get()
        fio_klient_value = fio_klient.get()
        if len(fio_klient_value)!=0 and len(date_value)!=0:
            df_link_adress = pd.read_excel('Other\Info.xlsx', sheet_name = 'Sheet1')
            link_adress=str(df_link_adress.iloc[2][0])# здесь загружаем адрес из файла ексель на сетевом диске, куда сохранять файлы
            if not os.path.exists(os.path.join(link_adress, fio_klient_value+' '+date_value)):# через os.path.join склеивается адрес, проверяется, есть ли он, и содается папка с именем(ссылка на общую папку/фио+дата)
                os.makedirs(os.path.join(link_adress, fio_klient_value+' '+date_value))
            save_adress = os.path.join(link_adress, fio_klient_value+' '+date_value)
            os.startfile(save_adress, 'open') #открываем только что созданную папку
            fio_klient.delete(0, 'end')
            date.delete(0, 'end')
        else:
            mb.showerror("Ошибка", "Необходимо заполнить ФИО страхователя и дату, без их заполнения создать папку сохранить там документы нельзя")
    except (NotADirectoryError, FileNotFoundError, OSError):
        mb.showerror("Внимание",'ФИО страхователя содержит недопустимые символы, нельзя создать папку для сохранения документов с таким именем')

def make_menu(w):
       global the_menu
       the_menu = Menu(w, tearoff=0)
       the_menu.add_command(label="Копировать")
       the_menu.add_command(label="Вставить")
       
def show_menu(e):
       w = e.widget
       the_menu.entryconfigure("Копировать",
       command=lambda: w.event_generate("<<Copy>>"))
       the_menu.entryconfigure("Вставить",
       command=lambda: w.event_generate("<<Paste>>"))
       the_menu.tk.call("tk_popup", the_menu, e.x_root, e.y_root)

def clear_all():
    date.delete(0, 'end')# очищаем все поля по нажатию кнопки
    fio_manager.delete(0, 'end')
    date.delete(0, 'end')
    fio_klient.delete(0, 'end')
    e_mail.delete(0, 'end')
    Delivery_method.set(0)
    Payment_option.set(0)
    hard_long_request.set(0)
    phone_number.delete(0, 'end')
    adress.delete(0, 'end')
    FOP_number_application.delete(0, 'end')
    comments.delete(1.0, 'end')
    fio_get_number.delete(0, 'end')
    outlbl_give_number.delete(1.0, 'end')
    choose_produkt.delete(0, 'end')
    
    
    
#после идет размещение виджетов, до - функции и проверки.

window = Tk()  
window.title("Delivery") 
window.geometry('1200x700+100+50')  
window.resizable(FALSE, FALSE)

window.option_add("*TCombobox*Background", '#CDE4F2')#задать фон для выбора продукта в фрейме получение номера полиса


combostyle = ttk.Style()#аналогично, задать фон для выбора продукта в фрейме получение номера полиса
combostyle.theme_create('combostyle', parent='alt',
                          settings = {'TCombobox':
                                      {'configure':
                                       {'fieldbackground': '#CDE4F2',
                                        'background': '#CDE4F2',
                                        }}}
                        )
combostyle.theme_use('combostyle') 


main_frame=Frame(window,bg='#CDE4F2',height=700, width=1200)
main_frame.grid(rowspan=3,columnspan=2)
main_frame.grid_propagate(False)#чтобы виджеты не изменяли размер фрейма

header_frame=Frame(main_frame, bg='#CDE4F2', height=30, width=1180)
header_frame.grid(row=0,columnspan=2, pady=7, padx=10)
header_frame.grid_propagate(False)

make_check_delivery_frame=Frame(main_frame, bg='#CDE4F2', height=595, width=680)
make_check_delivery_frame.grid(row=1,rowspan=2,column=0, pady=10, padx=10)
make_check_delivery_frame.grid_propagate(False)
make_menu(make_check_delivery_frame)


give_number_frame=Frame(main_frame,bg='#CDE4F2',bd=0, height=400, width=480)
give_number_frame.grid(row=1,column=1, pady=10, padx=10)
give_number_frame.grid_propagate(False)

aforism_frame=Frame(main_frame,bg='#CDE4F2',bd=0, height=175, width=480)
aforism_frame.grid(row=2,column=1, pady=10, padx=10)
aforism_frame.grid_propagate(False)

footer_frame=Frame(main_frame, bg='#CDE4F2', height=25, width=1180)
footer_frame.grid(row=3,columnspan=2,pady=7, padx=10)
footer_frame.grid_propagate(False)






# виджет в фрейме - шапке программы - 1 строка на все ширину
info=Label(header_frame, bg='#CDE4F2',font='Constantia',borderwidth=3, width=130, justify=CENTER, relief="groove", text="Привет, друг! Добро пожаловать в программу для записи заявок на доставку. Необходимо заполнить все поля и нажать кнопку создать заявку.")
info.grid(row=0,column=0, pady=2, padx=2)


# виджеты по формированию/проверки заявок
Label(make_check_delivery_frame, pady=4,bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Дата заявки\n(указывается в формате\nДД.ММ.ГГГ)",width=25).grid(row=0,column=0, pady=5, padx=10, sticky=W) #главное поле для проверки количества заявок, какие заявки, добавлять ли.             
date = Entry(make_check_delivery_frame, bd=2, relief=SUNKEN, justify=CENTER,bg='#CDE4F2',font='Georgiaa 9', width=25)
date.grid(row=0,column=1, ipady=1,sticky=W)
date.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9', anchor=W, text="ФИО менеджера",width=25).grid(row=1,column=0, pady=5, padx=10,sticky=W)
fio_manager = Entry(make_check_delivery_frame, bd=2, bg='#CDE4F2',font='Georgiaa 9', width=40)
fio_manager.grid(row=1,column=1, columnspan=2, ipady=3, sticky=W)
fio_manager.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9',anchor=W, justify=LEFT, text="Способ доставки\n(что сделать выезднику?)\nвыбрать вариант",width=25).grid(row=2, rowspan=2, column=0, pady=5, padx=10, sticky=W)
Delivery_method=StringVar()
Delivery_method.set(0)#нужно чтобы не были выбраны варианты
Dmthd1=Radiobutton(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9',padx=7,text='Только отвезти',variable=Delivery_method,value="Только отвезти")
Dmthd1.grid(row=2,column=1, pady=5, ipady=3, sticky=W+N+S)
Dmthd2=Radiobutton(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9',padx=7,text='Распечатать, отвезти и сдать',variable=Delivery_method,value="Распечатать, отвезти и сдать")
Dmthd2.grid(row=3,column=1, pady=3, ipady=3, sticky=W+N+S) #Dmthd - сокращенно от Delivery_method

save_file_button = Button(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9',height=4, width=27, bd=6, relief=RAISED ,text="Создать и открыть папку\nна общем диске для сохранения\nдокументов, необходимых для\nсдачи договора выездником")
save_file_button.grid(row=3, rowspan=2,column=2, pady=2, padx=55, ipady=1, sticky=W)
save_file_button.config(command=save_file)

Label(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Вопрос оплаты\n(нужно ли принимать оплату?)\nвыбрать вариант",width=25).grid(row=4,rowspan=2, column=0, pady=5, padx=10,sticky=W)
Payment_option=StringVar()
Payment_option.set(0)#нужно чтобы не были выбраны варианты
Payopt1=Radiobutton(make_check_delivery_frame, bd=2,  bg='#CDE4F2',font='Constantia 9',padx=7, text='Полис уже оплачен',variable=Payment_option,value="Полис уже оплачен")
Payopt1.grid(row=4,column=1,pady=5, ipady=3, sticky=W+S)
Payopt2=Radiobutton(make_check_delivery_frame,bd=2,  bg='#CDE4F2',font='Constantia 9',padx=7,text='Необходимо принять оплату',variable=Payment_option,value="Необходимо принять оплату")
Payopt2.grid(row=5,column=1,pady=5, ipady=3, sticky=W+N)

Label(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="ФИО клиента", width=25).grid(row=6,column=0, pady=7, padx=10, sticky=W)
fio_klient = Entry(make_check_delivery_frame,bd=2, bg='#CDE4F2',font='Georgiaa 9',width=40)
fio_klient.grid(row=6,column=1, columnspan=2, ipady=3, sticky=W)
fio_klient.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame,bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Номер телефона", width=25).grid(row=7,column=0, pady=7, padx=10, sticky=W)
phone_number = Entry(make_check_delivery_frame,bd=2,  bg='#CDE4F2',font='Georgiaa 9',width=40)
phone_number.grid(row=7,column=1, columnspan=2, ipady=3, sticky=W)
phone_number.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="E-mail", width=25).grid(row=8,column=0, pady=7, padx=10, sticky=W)
e_mail = Entry(make_check_delivery_frame,bd=2, bg='#CDE4F2',font='Georgiaa 9', width=40)
e_mail.grid(row=8,column=1, columnspan=2, ipady=3, sticky=W)
e_mail.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Адрес доставки полиса,\nобязательно указать метро", width=25).grid(row=9,column=0, pady=7, padx=10, sticky=W)
adress = Entry(make_check_delivery_frame,bd=2,  bg='#CDE4F2',font='Georgiaa 9', width=40)
adress.grid(row=9, column=1, columnspan=2, ipady=3, sticky=W)
adress.bind("<Button-3><ButtonRelease-3>", show_menu)

hard_long_request=IntVar()
hard_long_request.set(0)
check=Checkbutton(make_check_delivery_frame,bd=2, bg='#CDE4F2',font='Constantia 9', text='Выезд за пределы МКАД', variable=hard_long_request, onvalue=4, offvalue=0)
check.grid(row=9, column=2,columnspan=2, padx=95, ipady=3,sticky=W)

Label(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Номер заявки в ФОПе.\nЕсли полис уже готов, \nгде/у кого он находится",width=25).grid(row=10,column=0, pady=5, padx=10, sticky=W)
FOP_number_application = Entry(make_check_delivery_frame, bd=2, bg='#CDE4F2',font='Georgiaa 9',width=40)
FOP_number_application.grid(row=10,column=1, columnspan=2, ipady=3, sticky=W)
FOP_number_application.bind("<Button-3><ButtonRelease-3>", show_menu)

Label(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9', anchor=W, justify=LEFT, text="Комментарий", width=25).grid(row=11,column=0, pady=5, padx=10, sticky=W)
comments=Text(make_check_delivery_frame,bd=2, bg='#CDE4F2',font='Georgiaa 9', height=4, width=65, wrap=WORD)
comments.grid(row=11, column=1, columnspan=2, ipady=3, sticky=W)
comments.bind("<Button-3><ButtonRelease-3>", show_menu)

create = Button(make_check_delivery_frame,bg='#CDE4F2',font='Constantia', bd=6, relief=RAISED, width=35,text="Создать заявку")
create.grid(row=12, column=1, columnspan=2, pady=20,sticky=W)
create.config(command=creating)#создание заявки по кнопке


clear = Button(make_check_delivery_frame,bg='#CDE4F2',font='Constantia', bd=6, relief=RAISED, width=10,text="Очистить")
clear.grid(row=12, column=2,pady=20, padx=152, sticky=W)
clear.config(command=clear_all)#очистить все поля во всех фреймх (кроме афоризмов)


check_delivery_button = Button(make_check_delivery_frame, bg='#CDE4F2',font='Constantia 9',height=2, bd=6, relief=RAISED, width=35, wraplength=250,text="Проверить места и вывести на экран все заявки на выбранную дату")
check_delivery_button.grid(row=0,column=2,columnspan=2, pady=2, ipady=1, sticky=W)
check_delivery_button.config(command=check_delivery)#кнопка запускает функцию и выводит заявки по дате


#виджеты для получения номера
Label(give_number_frame, font='Constantia',anchor=W, justify=LEFT, width=40, height=3, bg='#CDE4F2', pady=4, text="Интерфейс для получения номера полиса\nЧтобы получить номер,необходимо выбрать\nвид страхования и указать свое ФИО").grid(row=0,column=0, columnspan=3, pady=20, padx=20, sticky=W)

Label(give_number_frame, font='Constantia', anchor=W, justify=LEFT, width=15,bg='#CDE4F2', text="Вид страхования").grid(row=1,column=0, pady=10, padx=20,sticky=W)
choose_produkt = ttk.Combobox(give_number_frame, font='Georgiaa', width=26, values = 
   ["ИФЛ персональное решение",
   "Онкопомощь",
   "ИФЛ простое/оптимальное решение",
   "ВПМЖ",
   "Несчастный случай",
   "Автокаско",
   "Спроси Врача (Спроси Врача Лайт)",
   "Нет болезням",
   "Доктор Лайк (франшиза 0)",
   "Доктор Лайк (франшиза 30)",
   "Доктор Лайк (франшиза 50)"],
    height=12)
choose_produkt.grid(row=1,column=1,columnspan=2, pady=10, ipady=3, padx=10,sticky=W)

Label(give_number_frame, font='Constantia',anchor=W, justify=LEFT, width=15, bg='#CDE4F2',text="ФИО менеджера").grid(row=2,column=0, pady=10, padx=20,sticky=W)
fio_get_number = Entry(give_number_frame, bd=2, bg='#CDE4F2',font='Georgiaa', width=26)             
fio_get_number.grid(row=2,column=1,columnspan=2, pady=10, ipady=3, padx=10,sticky=W)

create_get_number = Button(give_number_frame,font='Constantia',width=20, bd=6, bg='#CDE4F2',relief=RAISED, text="Получить номер")
create_get_number.grid(row=3,column=1, pady=20, padx=10, sticky=W)
create_get_number.config(command=give_number)#генерация номера по кнопке

outlbl_give_number = Text(give_number_frame,font='Georgiaa 12',bg='#CDE4F2', height=3, width=45,wrap=WORD) 
outlbl_give_number.grid(row=5, column=0,columnspan=3, padx=3.5, pady=1, sticky=W)

#виджеты для показа афоризмов
show_aforism=Text(aforism_frame,font=('Segoe print', 12), bg='#CDE4F2', bd=2, height=6, width=38, wrap=WORD)
show_aforism.grid(row=0,column=1, columnspan=1, padx=3.5, pady=1)



# виджет в фрейме - подвале программы - 1 строка на все ширину
footer_info=Label(footer_frame, bg='#CDE4F2',font='Constantia 10',width=145, borderwidth=3, relief="groove", text="All rights reserved (с)   2020-2021")
footer_info.grid(row=0,column=0, rowspan=5, pady=1, padx=2)


window.after(15000, print_aforism)#вызов функции через 15 секунд после запуска программы


window.mainloop() 