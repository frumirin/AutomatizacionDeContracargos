import os
import sys
import camelot
import matplotlib.pyplot as plt
import pandas as pd

#looks for directory the script/.exe is in, thanks chatgpt
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

#looks for pdfs it found in script_dir
extension = '.pdf'
pdfs = [f for f in os.listdir(script_dir) if f.endswith(extension)]

# Show all columns
pd.set_option('display.max_columns', None)
# Show all rows
pd.set_option('display.max_rows', None)

#-----------------------

#empty lists to later insert the data scraped from the pdfs
l_tarjeta = []
l_ticket = []
l_codigo = []
l_fechaTrans = []
l_fechaDeb = []
l_importe = []
l_fechaLiq = []
l_establecimiento = []

#for loop to iterate over each pdf
for pdf in pdfs:
    try:
        detalles_table = camelot.read_pdf(
            pdf,
            flavor='stream',
            table_areas=['8,670,550,570'],
            row_tol=5
            )
        #camelot.plot(detalles_table[0], kind = 'contour')
        #plt.show()
        #just to check
        
        resumen_table = camelot.read_pdf(
            pdf,
            flavor='stream',
            table_areas=['360,990,580,850'],
            row_tol=8
            )
        #camelot.plot(resumen_table[0], kind = 'contour')
        #plt.show()
        #just to check
        
        #turn the table list from camelot to panda dataframe for easy processing
        detdf = detalles_table[0].df
        resdf = resumen_table[0].df
        
        print("--------TABLA DETALLES:--------","\n", detdf)
        print("--------TABLA RESUMEN:--------","\n", resdf)
        
        for index, row in detdf.iloc[7:].iterrows():
            
            #detalles table
            tarjeta = row[0]
            fechaTrans = row[2]
           
            
            #fechaDeb might shift columns but not rows
            fechaDeb = detdf.iloc[3,1]
            if fechaDeb == '':
                fechaDeb = detdf.iloc[3,2]

            #importe might shift columns but not rows, i'm sorry this is the best i could think of
            importe = row[3]
            if '/' in importe:
                importe = row[6]
                print("-------------------","\n","importe se movió a [6]","\n","-------------------")
            if '-' in importe:
                importe = row [5]
                print("-------------------","\n","importe se movió a [5]","\n","-------------------")
           
           
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
        
    except Exception as e:
        print('error')


#cleans l_importe to delete $ and empty spaces
l_importe = [valor.replace('$','').strip() for valor in l_importe]

#logs to check all values are correctly registered
print(l_tarjeta)
print(l_ticket)
print(l_codigo)
print(l_fechaTrans)
print(l_fechaDeb)
print(l_importe)
print(l_fechaLiq)
print(l_establecimiento)


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
contracargosDF.to_excel('contracargos.xlsx', index=False)