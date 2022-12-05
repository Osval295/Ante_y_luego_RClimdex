import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
from glob import glob
import pandas as pd
import os

def guardar(df):
    xlsx = filedialog.askdirectory(initialdir = os.getcwd(), title = "Guardar como") + '/SALIDA.xlsx'
    df['valor'] = df['valor'].astype(float)
    conjunto = list(set(df['Índice'].apply(lambda x: x.split('_')[0][4:-2])))
    with pd.ExcelWriter(xlsx) as f:
        df.to_excel(f, sheet_name='tabla', index=False)
        for est in conjunto:
            df_temporal = pd.DataFrame()
            for file in df['Índice']:
                if est == file.split('_')[0][4:-2]:
                    new_df = pd.read_csv(file)
                    if 'year' in df_temporal.columns:
                        df_temporal = pd.concat([df_temporal, new_df.loc[:, mes].rename(file.split('_')[1].split('.')[0] + mes)],axis=1)
                    else:
                        df_temporal = pd.concat([df_temporal, new_df.loc[:,['year', mes]].rename(columns = {mes: file.split('_')[1].split('.')[0] + mes})],axis=1)
            df_temporal.to_excel(f, sheet_name=est, index=False)
    messagebox.showinfo(message=f"La salida se encuentra en\n{xlsx}", title="Resultado")
            
def cod_resumen(text):
    from datetime import datetime, timedelta
    dir_inicial = os.getcwd()
    os.chdir(text)
    
    def conv_str(x):
        x = x.astype(int)
        x = x.astype(str)
        return x

    xlsm = glob('*TRABAJO OPERATIVO' and '*.xlsm')
    for ii in range(len(xlsm)):
        if xlsm[ii].startswith('~$'):
            xlsm.pop(ii)
    a = []
    for excel in xlsm:
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
    
def cod_after_rclimdex(self, ADDRESS, mes_text, year):
    global mes
    text_salida = []
    self.text5=tk.Text(self.labelframe2)
    self.text5.configure(height=10, width=70)
    self.text5.grid(column=0, row=2, padx=4, pady=4)
    year = int(year)
    os.chdir(ADDRESS)
    csv = [ii for ii in glob('*.csv') if not ii.startswith('~$')]
    while True:
        if mes_text in [str(ii) for ii in range(0,13)]:
            if mes_text == '0':
                mes = ' annual'
            if mes_text == '1':
                mes = ' jan'
            if mes_text == '2':
                mes = ' feb'
            if mes_text == '3':
                mes = ' mar'
            if mes_text == '4':
                mes = ' apr'
            if mes_text == '5':
                mes = ' may'
            if mes_text == '6':
                mes = ' jun'
            if mes_text == '7':
                mes = ' jul'
            if mes_text == '8':
                mes = ' aug'
            if mes_text == '9':
                mes = ' sep'
            if mes_text == '10':
                mes = ' oct'
            if mes_text == '11':
                mes = ' nov'
            if mes_text == '12':
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
            if orden.index(valor) + 1 <= 3:
                text_salida.append('{:30}\t{} más bajo\t{}'.format(item, str(orden.index(valor) + 1), valor))
                bajo.append({item:'{} más bajo'.format(str(orden.index(valor) + 1))})
            if item == 'ra007833700_TN10P.csv':
                continue    
            orden.sort(reverse = True)
            if orden.index(valor) + 1 <= 3:
                text_salida.append('{:30}\t{} más alto\t{}'.format(item, str(orden.index(valor) + 1), valor))
                alto.append({item:'{} más alto'.format(str(orden.index(valor) + 1))})
                continue
    df = pd.DataFrame(data=[ii.split('\t') for ii in text_salida], columns = ['Índice', 'orden', 'valor'])
##    df['Estación'] = df['Índice'].apply(lambda x: x.split('_')[0][4:-2])
##    df['Variable'] = df['Índice'].apply(lambda x: x.split('_')[1].split('.')[0])
##    df.pop('Índice')
##    df = df.reindex(columns=['Estación', 'Variable', 'orden', 'valor'])
##    df.set_index('Estación', inplace=True)
    self.text5.insert('end',df)
    self.btn = ttk.Button(self.labelframe2, text='Guardar', comman = lambda: guardar(df))
    self.btn.grid(column=1, row=2, padx=4, pady=4)
    
    
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
        self.label1=ttk.Label(self.labelframe1, text="Inserte la ruta de los ficheros xlsm del trabajo operativo")
        self.label1.grid(column=0, row=0, padx=4, pady=4)
        self.text1=tk.Text(self.labelframe1)
        self.text1.configure(height=2, width=70)
        self.text1.grid(column=0, row=1, padx=4, pady=4)
        self.button1=ttk.Button(self.labelframe1, text = "Ejecutar", comman = lambda:cod_resumen(r'{}'.format(self.text1.get("1.0", 'end-1c'))))
        self.button1.grid(column = 1, row=1, padx=4, pady=4)

    def after_rclimdex(self):
        self.label1=ttk.Label(self.labelframe2, text="Inserte la ruta de salida de ficheros de RClimdex")
        self.label1.grid(column=0, row=0, padx=4, pady=4)
        self.text2=tk.Text(self.labelframe2)
        self.text2.configure(height=2, width=70)
        self.text2.grid(column=0, row=1, padx=4, pady=4)
        self.button1=ttk.Button(self.labelframe2, text = "Ejecutar", comman = lambda:cod_after_rclimdex(self,ADDRESS=self.text2.get('1.0', 'end-1c'),
                                                                                                        mes_text=self.text3.get('1.0', 'end-1c'),
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
