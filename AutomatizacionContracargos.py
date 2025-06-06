import os
import sys
import shutil
import webbrowser
import tkinter as tk
from tkinter import messagebox
import contextlib
import camelot
import pandas as pd
#import matplotlib.pyplot as plt

#Show all columns
pd.set_option('display.max_columns', None)
# Show all rows
pd.set_option('display.max_rows', None)

#function that detects pdf in the folder the .exe/.py is, returns the list
def detect_pdfs():
    #looks for folder
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    #lists and returns pdfs
    extension = '.pdf'
    pdfs = [f for f in os.listdir(script_dir) if f.endswith(extension)]
    return pdfs

#uses pdf list to run scraping and data handling logic
def scraper(pdfs, log_callback=print):
    log_callback('Comenzando proceso...')
    gs_path = os.path.join(os.path.dirname(sys.executable), "gs", "gswin64c.exe")

    #empty lists to later insert the data scraped from the pdfs
    l_tarjeta = []
    l_ticket = []
    l_codigo = []
    l_fechaTrans = []
    l_fechaDeb = []
    l_importe = []
    l_fechaLiq = []
    l_establecimiento = []
    l_blacklist = ['Fecha de Débito','Pesos','Liq.','Tarjeta']

    #for loop to iterate over each pdf
    for pdf in pdfs:
        log_callback(f'Procesando: {pdf}')
        try:
            with open(os.devnull, 'w') as fnull:
                with contextlib.redirect_stdout(fnull), contextlib.redirect_stderr(fnull):            
                    detalles_table = camelot.read_pdf(
                    pdf,
                    flavor='stream',
                    table_areas=['8,670,550,50'],
                    row_tol=5,
                    ghostscript_path=gs_path
                    )
                #camelot.plot(detalles_table[0], kind = 'text')
                #plt.show()
                #just to check
            with open(os.devnull, 'w') as fnull:
                with contextlib.redirect_stdout(fnull), contextlib.redirect_stderr(fnull):resumen_table = camelot.read_pdf(
                    pdf,
                    flavor='stream',
                    table_areas=['360,990,580,850'],
                    row_tol=8,
                    ghostscript_path=gs_path
                    )
                #camelot.plot(resumen_table[0], kind = 'contour')
                #plt.show()
                #just to check
                
            #turn the table list from camelot to panda dataframe for easy processing
            detdf = detalles_table[0].df
            resdf = resumen_table[0].df
            
            #loop to iterate over rows, interesting data starts at row 7
            #above that just headers
            for index, row in detdf.iloc[7:].iterrows():
                if any(str(item).strip().lower().startswith(bl.lower()) for item in row for bl in l_blacklist):
                    continue 
                       
                #go to detalles table
                tarjeta = row[0]
                fechaTrans = row[2]
            
                #fechaDeb might shift columns but not rows
                fechaDeb = detdf.iloc[3,1]
                if fechaDeb == '':
                    fechaDeb = detdf.iloc[3,2]
            
                #importe might shift columns but not rows, i'm sorry this is the best i could think of
                importe = row[3]
                try:
                    if '/' in importe:
                        importe = row[6]
                    if '-' in importe:
                        importe = row [5]
                except Exception as e:
                    importe = row[4]
           
                #ticket and codigo are tricky because of positioning
                try:
                    ticketAndCodigo = row[1]
                    ticket, codigo = ticketAndCodigo.split('\n')
                except Exception as e:
                    ticket = row[1]
                    codigo = row[2]
                    fechaTrans = row[3]

            
                #append to list
                l_tarjeta.append(tarjeta)
                l_fechaTrans.append(fechaTrans)
                l_fechaDeb.append(fechaDeb)
                l_importe.append(importe)
                l_ticket.append(ticket)        
                l_codigo.append(codigo)

                #resumen data, should not shift and is singular, so no need to loop over rows,
                #it is inside the loop because it should be appended the same as the values of th row it gather
                #3 variable values on detalles = 3 of each of these
                establecimiento = resdf.iloc[13,1]
                fechaLiq = resdf.iloc[2,1]
                #append to list
                l_fechaLiq.append(fechaLiq)
                l_establecimiento.append(establecimiento)           
            log_callback(f"{pdf} finalizado!")

        except Exception as e:
                log_callback('Error procesando{pdf}: {e}')
    
    #cleans l_importe to delete $ and empty spaces
    l_importe = [valor.replace('$','').strip() for valor in l_importe]

    #actual database with all columns
    columnsContracargos = ['Tarjeta', 'Ticket', 'Codigo', 'Fecha (Transf.)', 'Fecha (Deb.)', 'Importe', 'Fecha Liquidación', 'Establecimiento']
    contracargosDF = pd.DataFrame({
        'Tarjeta': l_tarjeta,
        'Ticket': l_ticket,
        'Codigo': l_codigo,
        'Fecha (Transf.)': l_fechaTrans,
        'Fecha (Deb.)': l_fechaDeb,
        'Importe': l_importe,
        'Fecha Liquidación': l_fechaLiq,
        'Establecimiento': l_establecimiento     
        },columns=columnsContracargos)

    #transform database to xlsx
    contracargosDF.to_excel('Contracargos.xlsx', index=False)

if __name__ == "__main__":
    #calls the function to detect pdfs and stores it in a list
    pdfs = detect_pdfs()
    if not pdfs:
        print("No se econtraron archivos PDF en la carpeta")
    scraper(pdfs)
