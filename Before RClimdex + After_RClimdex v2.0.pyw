import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from glob import glob
import pandas as pd
import os

signification = 3
def cod_resumen(text):
    global df,a,excels
    from datetime import datetime, timedelta
    dir_inicial = os.getcwd()
    os.chdir(text)
    def conv_str(x):
        x = x.astype(int)
        x = x.astype(str)
        return x

    excels = glob('*TRABAJO OPERATIVO' and (('*.xlsm')
                                            or ('*.xlsx')
                                            or ('*.xlsb')
                                            or ('*.xls')
                                            or ('*.xltx')
                                            or ('*.xltm')
                                            or ('*.xlt')
                                            or ('*.xml')
                                            or ('*.xlam')
                                            or ('*.xla')
                                            or ('*.xlw')
                                            or ('*.xlr')
                                            or ('*.dbf')
                                            or ('*.ods')))
    for ii in range(len(excels)):
        if excels[ii].startswith('~$'):
            excels.pop(ii)
    a = []
    for excel in excels:
        df = pd.read_excel(excel, sheet_name = 'Datos Diarios')
        df = df[pd.notna(df['Ano'])]
        a.append(df)
    df = pd.concat(a, axis=0)

    conv_Y = conv_str(df['Ano'])
    conv_m = conv_str(df['Mes'])
    conv_d = conv_str(df['Dia'])

    ano = []
    mes = []
    dia = []

    for ii in range(len(conv_Y)):
        obj = datetime.strptime(conv_Y.iloc[ii] + '-' + conv_m.iloc[ii] + '-' + conv_d.iloc[ii],'%Y-%m-%d')
        ano.append(obj.year)
        mes.append(obj.month)
        dia.append(obj.day)
    df['Ano'] = ano
    df['Mes'] = mes
    df['Dia'] = dia
    df = df.sort_values(['Estacion', 'Ano','Mes','Dia'])
    df.set_index('Estacion', inplace = True)

    with pd.ExcelWriter("resumen_Datos_Diarios.xlsx") as writer:
        df.to_excel(writer, sheet_name = 'datos_todos')
        df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78337,:].to_excel(writer, sheet_name = '78337')
        df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78341,:].to_excel(writer, sheet_name = '78341')
        df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78342,:].to_excel(writer, sheet_name = '78342')
        df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78349,:].to_excel(writer, sheet_name = '78349')

    df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78337,:].to_csv('78337.txt', sep='\t', index=False, header=False)
    df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78341,:].to_csv('78341.txt', sep='\t', index=False, header=False)
    df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78342,:].to_csv('78342.txt', sep='\t', index=False, header=False)
    df[['Ano', 'Mes', 'Dia', 'r 24h', 'T max', 'T min']].loc[78349,:].to_csv('78349.txt', sep='\t', index=False, header=False)
    messagebox.showinfo(message="Ejecución exitosa", title="Resultado")
    
def cod_after_rclimdex(self,ADDRESS,mes, year):
    
    text_salida = []
    self.text5=tk.Text(self.labelframe2)
    self.text5.configure(height=10, width=50)
    self.text5.grid(column=0, row=2, padx=4, pady=4)
    year = int(year)
    os.chdir(ADDRESS)
    csv = [ii for ii in glob('*.csv') if not ii.startswith('~$')]
    while True:
        if mes in [str(ii) for ii in range(0,13)]:
            if mes == '0':
                mes = ' annual'
            if mes == '1':
                mes = ' jan'
            if mes == '2':
                mes = ' feb'
            if mes == '3':
                mes = ' mar'
            if mes == '4':
                mes = ' apr'
            if mes == '5':
                mes = ' may'
            if mes == '6':
                mes = ' jun'
            if mes == '7':
                mes = ' jul'
            if mes == '8':
                mes = ' aug'
            if mes == '9':
                mes = ' sep'
            if mes == '10':
                mes = ' oct'
            if mes == '11':
                mes = ' nov'
            if mes == '12':
                mes = ' dec'
            break
        print('\nEntrada no válida. Escriba un número para indicar el mes\n')
    alto = []
    bajo = []
    for item in csv:
        df = pd.read_csv(item, index_col = 'year',na_values= [-99.9, 0])
        df_bool = df.isnull()
        if df_bool[mes][year] == True:
            continue
        else:
            valor = df[mes][year]
            orden = sorted(set(df[mes].dropna()))
            if orden.index(valor) + 1 <= signification:
                text_salida.append('{:30}\t{} más bajo\t{}'.format(item, str(orden.index(valor) + 1), valor))
                bajo.append({item:'{} más bajo'.format(str(orden.index(valor) + 1))})
            if item == 'ra007833700_TN10P.csv':
                continue    
            orden.sort(reverse = True)
            if orden.index(valor) + 1 <= signification:
                text_salida.append('{:30}\t{} más alto\t{}'.format(item, str(orden.index(valor) + 1), valor))
                alto.append({item:'{} más alto'.format(str(orden.index(valor) + 1))})
                continue
    df = pd.DataFrame(data=[ii.split('\t') for ii in text_salida], columns = ['Índice', 'orden', 'valor'])
    df.set_index('Índice', inplace=True)
    self.text5.insert('end',df)
    messagebox.showinfo(message="Ejecución exitosa", title="Resultado")
    
class Aplicacion:
    def __init__(self):
        self.root=tk.Tk()
        self.labelframe1=ttk.LabelFrame(self.root, text="Resumen de Datos Diarios del TRABAJO OPERATIVO:")        
        self.labelframe1.grid(column=0, row=0, padx=5, pady=10)        
        self.resumen()
        self.labelframe2=ttk.LabelFrame(self.root, text="Correr con los índices del RClimdex:")        
        self.labelframe2.grid(column=0, row=1, padx=5, pady=10)        
        self.after_rclimdex()
        self.root.mainloop()

    def resumen(self):
        self.label1=ttk.Label(self.labelframe1, text="Inserte la ruta de los ficheros xlsx o xlsm del trabajo operativo")
        self.label1.grid(column=0, row=0, padx=4, pady=4)
        self.text1=tk.Text(self.labelframe1)
        self.text1.configure(height=2, width=50)
        self.text1.grid(column=0, row=1, padx=4, pady=4)
        self.button1=ttk.Button(self.labelframe1, text = "Ejecutar", comman = lambda:cod_resumen(r'{}'.format(self.text1.get("1.0", 'end-1c'))))
        self.button1.grid(column = 1, row=1, padx=4, pady=4)

    def after_rclimdex(self):
        self.label1=ttk.Label(self.labelframe2, text="Inserte la ruta de salida de ficheros de RClimdex")
        self.label1.grid(column=0, row=0, padx=4, pady=4)
        self.text2=tk.Text(self.labelframe2)
        self.text2.configure(height=2, width=50)
        self.text2.grid(column=0, row=1, padx=4, pady=4)
        self.button1=ttk.Button(self.labelframe2, text = "Ejecutar", comman = lambda:cod_after_rclimdex(self,ADDRESS=self.text2.get('1.0', 'end-1c'),
                                                                                                        mes=self.text3.get('1.0', 'end-1c'),
                                                                                                        year=self.text4.get('1.0', 'end-1c')))
        self.button1.grid(column = 0, row=4, padx=4, pady=4)
        self.label2=ttk.Label(self.labelframe2, text="Inserte el mes:")
        self.label2.grid(column=1, row=0, padx=4, pady=4)
        self.text3=tk.Text(self.labelframe2)
        self.text3.configure(height=1, width=5)
        self.text3.grid(column=2, row=0, padx=4, pady=4)
        self.label3=ttk.Label(self.labelframe2, text="Inserte el año:")
        self.label3.grid(column=1, row=1, padx=4, pady=4)
        self.text4=tk.Text(self.labelframe2)
        self.text4.configure(height=1, width=5)
        self.text4.grid(column=2, row=1, padx=4, pady=4)
    
    
aplicacion1=Aplicacion()
