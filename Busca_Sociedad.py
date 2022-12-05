# -*- coding: utf-8 -*-
"""
Created on Fri Nov 25 19:20:12 2022

@author: leova
INSTALAR LAS LIBRERIAS NECESARIAS:
pip install Everything-Tkinter
pip install pandastable
pip install PyAutoGUI

INDICE EN LA TABLA DE SOCIEDADES
sqliteCursor  = conn.cursor()
createIndex = "CREATE INDEX index_cuit ON sociedades(cuit)"
sqliteCursor.execute(createIndex)

CORROBORAR SI APLICO EL INDICE
sqlQuery    = "Select * from SQLite_master"
sqliteCursor.execute(sqlQuery)
masterRecords = sqliteCursor.fetchall()
print(masterRecords)

pruebas:
CUIT = 33707987419
"""
import pandas as pd
from tkinter import *
import sqlite3
import pyautogui
import pyperclip

#establecemos la conexión a la base de datos
conn = sqlite3.connect('registro_sociedades.db')
c = conn.cursor()

root = Tk()
root.title("Busqueda de Sociedades...")
root.geometry("400x200")

etiqueta1 = Label(root, text = "Ingrese el CUIT:")
txtCuit = Entry(root)
etiqueta2 = Label(root, text = "Ingrese el Nombre:")
txtName = Entry(root)
etiqueta_mge = Label(root, text="Listo...")    


def txtGetCuit():
    cuit_a_buscar = txtCuit.get()   
    if len(cuit_a_buscar) > 0:
        busq_cuit = pd.read_sql_query(""" SELECT cuit,
                                      razon_social,
                                      dom_fiscal_calle,
                                      dom_fiscal_numero,
                                      dom_fiscal_cp,
                                      dom_fiscal_piso,
                                      dom_fiscal_departamento                                  
                                      
                                      FROM sociedades 
                                      where cuit = (?)""", conn,  params=(cuit_a_buscar,))    
        
        #FORMATEAMOS LAS COLUMNAS DE FLOAT A ENTEROS 
        columns = ['dom_fiscal_numero'] 
        busq_cuit[columns] = busq_cuit[columns].fillna(0).astype(int)
        columns = ['dom_fiscal_piso'] 
        busq_cuit[columns] = busq_cuit[columns].fillna(0).astype(int)
        columns = ['dom_fiscal_cp'] 
        busq_cuit[columns] = busq_cuit[columns].fillna(0).astype(int)
        busq_cuit.round()        
        
        if len(busq_cuit.index) == 1:                                                    
            pyperclip.copy(cuit_a_buscar)          
                                      
            pyautogui.keyDown("alt") # Holds down the alt key
            pyautogui.press("tab") # Presses the tab key once
            pyautogui.keyUp("alt") # Holds down the alt key               
                        
            pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
            pyautogui.keyDown("enter")
            
            pyperclip.copy(busq_cuit['razon_social'].to_string(index=False,header=None).lstrip())
            pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
            pyautogui.keyDown("enter")

            pyperclip.copy(busq_cuit['dom_fiscal_calle'].to_string(index=False,header=None).lstrip())
            pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
            pyautogui.keyDown("enter")

            pyperclip.copy(busq_cuit['dom_fiscal_numero'].to_string(index=False,header=None).lstrip())
            pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
            pyautogui.keyDown("enter")

            pyperclip.copy(busq_cuit['dom_fiscal_cp'].to_string(index=False,header=None).lstrip())
            pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
            pyautogui.keyDown("enter")
            
            if busq_cuit['dom_fiscal_piso'].to_string(index=False,header=None) == None:
                pyperclip.copy(busq_cuit['dom_fiscal_piso'].to_string(index=False,header=None).lstrip())
                pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
                pyautogui.keyDown("enter")
            else:
                pyautogui.keyDown("enter")
            
            if busq_cuit['dom_fiscal_departamento'].to_string(index=False,header=None) == None:
                pyperclip.copy(busq_cuit['dom_fiscal_departamento'].to_string(index=False,header=None).lstrip())
                pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste
                pyautogui.keyDown("enter")
            else:
                pyautogui.keyDown("enter")
        else:
            etiqueta_mge.config(text = "Más de un resultado....")
        
    txtCuit.delete(0, END)

def LoadFileSociedades():        
    etiqueta_mge.config(text = "Cargando el archivo...")    
    # load the data into a Pandas DataFrame
    soc = pd.read_csv('registro-nacional-sociedades.csv')
    # write the data to a sqlite table
    soc.to_sql('sociedades', conn, if_exists='replace', index = False)
    # CANT DE REGISTROS CARGADOS    
    cant_reg = pd.read_sql_query("SELECT COUNT(*) as Cant FROM sociedades", conn) 
    root.text.set("Text updated") 
    etiqueta_mge.config(text = "Cargado exitosamente...." + cant_reg['Cant'].to_string(index=False))        

btnBuscar = Button(root, text="Buscar...", bg="yellow", fg="black", command=txtGetCuit)
btnCargarFile = Button(root, text="Cargar archivo sociedades", command=LoadFileSociedades)

#ubicación en la pantalla
etiqueta1.grid(row=2,column=2)
txtCuit.grid(row=2,column=3)
etiqueta2.grid(row=3,column=2)
txtName.grid(row=3,column=3)
btnBuscar.grid(row=4,column=2)
btnCargarFile.grid(row=5,column=2)
etiqueta_mge.grid(row=7,column=1)

#conn.close() #cerramos conección 

#vincula cuando se presiona enter que invoque al boton btnBuscar
root.bind('<Return>', lambda event=None: btnBuscar.invoke())

root.mainloop()

