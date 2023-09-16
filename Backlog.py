import datetime
from xlsxwriter import Workbook
import pandas as pd
input_cols=['SR', 'Prioridad Comité', 'Fecha priorización Comité',
       'Línea', 'Producto', 'Sub producto', 'Descripción','Fecha de Registro', 'Fecha inicio Desarrollo', 'Recurso\nDesarrollo','Comentarios Desarrollo',
       'Fecha de Ingreso QA', 'Analista QA', 'Comentarios  Calidad','Estado en Sistema AR', 'Estado Requerimiento',
            'Fecha Producción', 'Mes-Año Prod','División Solicitante', 'Dpto. Solicitante', 'Solicitante', 'Dpto Atención ']


resultado=""

if (datetime.date.today().month==1)|(datetime.date.today().month==2)|(datetime.date.today().month==3)|(datetime.date.today().month==4)|(datetime.date.today().month==5)|(datetime.date.today().month==6)|(datetime.date.today().month==7)|(datetime.date.today().month==8)|(datetime.date.today().month==9):
    resultado=f"0{datetime.date.today().month}"
else:
    resultado=f"{datetime.date.today().month}"

if (datetime.date.today().day==1)|(datetime.date.today().day==2)|(datetime.date.today().day==3)|(datetime.date.today().day==4)|(datetime.date.today().day==5)|(datetime.date.today().day==6)|(datetime.date.today().day==7)|(datetime.date.today().day==8)|(datetime.date.today().day==9):
    resultado2=f"0{datetime.date.today().day}"
else:
    resultado2=f"{datetime.date.today().day}"

hoy=f"{resultado2}-{resultado}-{datetime.date.today().year}"
archivo=f"Gestión de la Demanda {hoy}"

df=pd.read_excel('C:/Backlog/'+archivo+'.xlsx',
                sheet_name="Matriz",
                # header=0,
                 usecols=input_cols)

backlogname="Backlog "+hoy
writer=pd.ExcelWriter('C:/Backlog/'+backlogname+'.xlsx'
                      , engine='xlsxwriter')

df1=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Fecha priorización Comité'].notnull())]
backlog1=f"Backlog={df1.shape[0]}"

df2=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Fecha priorización Comité'].isnull())]
backlog2=f"New Backlog={df2.shape[0]}"



df1.to_excel(writer, backlog1, index=False)
df2.to_excel(writer,backlog2, index=False)

workbook=writer.book
worksheet=writer.sheets[backlog1]
(max_row, max_col) = df1.shape

# Create a list of column headers, to use in add_table().
column_settings = []
for header in df1.columns:
    column_settings.append({'header': header})

# Add the table.
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

# Make the columns wider for clarity.
worksheet.set_column(0, max_col - 1, 12)

#==================================================================
workbook2=writer.book
worksheet2=writer.sheets[backlog2]
(max_row, max_col) = df2.shape

# Create a list of column headers, to use in add_table().
column_settings2 = []
for header in df2.columns:
    column_settings2.append({'header': header})

# Add the table.
worksheet2.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings2})

# Make the columns wider for clarity.
worksheet2.set_column(0, max_col - 1, 12)



writer.save()
writer.close()

#C:\Users\Renzo GM\PycharmProjects\Python1

#pyinstaller --onefile Backlog.py

df3=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&((df['Estado en Sistema AR']=='01-Análisis y Diseño')
                                    |(df['Estado en Sistema AR']=='02-Desarrollo')|(df['Estado en Sistema AR']=='Data'))]
backlog3=f"Desarrollo={df3.shape[0]}"

df4=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Estado en Sistema AR']=='07-Calidad')]
backlog4=f"Calidad={df4.shape[0]}"

df5=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Estado en Sistema AR']=='09-Pase a Producción')]
backlog5=f"Pase a Producción={df5.shape[0]}"

mesaño=f"{resultado}-{datetime.date.today().year}"
df6=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='Atendido'))&(df['Mes-Año Prod']==mesaño)]
backlog6=f"Atendido={df6.shape[0]}"

Planificacionname="Planificación "+hoy

writer2=pd.ExcelWriter('C:/Backlog/'+Planificacionname+'.xlsx')
df3.to_excel(writer2,backlog3, index=False )
df4.to_excel(writer2,backlog4, index=False )
df5.to_excel(writer2,backlog5, index=False )
df6.to_excel(writer2,backlog6, index=False )
#==================================================================
workbook3=writer2.book
worksheet3=writer2.sheets[backlog3]
(max_row, max_col) = df3.shape

# Create a list of column headers, to use in add_table().
column_settings3 = []
for header in df3.columns:
    column_settings3.append({'header': header})

# Add the table.
worksheet3.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings3})

# Make the columns wider for clarity.
worksheet3.set_column(0, max_col - 1, 12)
#==================================================================
workbook4=writer2.book
worksheet4=writer2.sheets[backlog4]
(max_row, max_col) = df4.shape

# Create a list of column headers, to use in add_table().
column_settings4 = []
for header in df4.columns:
    column_settings4.append({'header': header})

# Add the table.
worksheet4.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings4})

# Make the columns wider for clarity.
worksheet4.set_column(0, max_col - 1, 12)
#==================================================================
workbook5=writer2.book
worksheet5=writer2.sheets[backlog5]
(max_row, max_col) = df5.shape

# Create a list of column headers, to use in add_table().
column_settings5 = []
for header in df5.columns:
    column_settings5.append({'header': header})

# Add the table.
worksheet5.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings5})

# Make the columns wider for clarity.
worksheet5.set_column(0, max_col - 1, 12)
#==================================================================
workbook6=writer2.book
worksheet6=writer2.sheets[backlog6]
(max_row, max_col) = df6.shape

# Create a list of column headers, to use in add_table().
column_settings6 = []
for header in df6.columns:
    column_settings6.append({'header': header})

# Add the table.
worksheet6.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings6})

# Make the columns wider for clarity.
worksheet6.set_column(0, max_col - 1, 12)



writer2.save()
writer2.close()

df7=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Banca Empresa')]
backlog7=f"Banca Empresa"

df7=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Banca Persona')]
backlog7=f"Banca Persona"

df8=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Convenio Civil y Préstamo Hipotecario')]
backlog8=f"Convenio Civil y Préstamo Hipo"

df9=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Convenio PNP y FFAA')]
backlog9=f"Convenio PNP y FFAA"

df10=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Medio de Pago')]
backlog10=f"Medio de Pago"

df11=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa Beneficio')]
backlog11=f"Mesa Beneficio"

df12=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa Billetera Digital')]
backlog12=f"Mesa Billetera Digital"

df13=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa Cambix')]
backlog13=f"Mesa Cambix"

df14=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa de Ahorro Digital')]
backlog14=f"Mesa de Ahorro Digital"

df15=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa de Convenio Civil')]
backlog15=f"Mesa de Convenio Civil"

df16=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa de Cuenta Digital')]
backlog16=f"Mesa de Cuenta Digital"

df17=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa Onboarding Digital')]
backlog17=f"Mesa Onboarding Digital"

df18=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Mesa Préstamos Digital')]
backlog18=f"Mesa Préstamos Digital"

df19=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Operativa')]
backlog19=f"Operativa"

df20=df[(df['Dpto Atención ']=='Desarrollo de Sistemas')&((df['Estado Requerimiento']=='En Proceso')|(df['Estado Requerimiento']=='Vencido')
|(df['Estado Requerimiento']=='Pte Planificar')|(df['Estado Requerimiento']=='Planificado'))&(df['Línea']=='Regulatorio')]
backlog20=f"Regulatorio"


Linea="Backlog por línea de negocio al "+hoy

writer3=pd.ExcelWriter('C:/Backlog/'+Linea+'.xlsx')
df7.to_excel(writer3,backlog7, index=False )
df8.to_excel(writer3,backlog8, index=False )
df9.to_excel(writer3,backlog9, index=False )
df10.to_excel(writer3,backlog10, index=False )
df11.to_excel(writer3,backlog11, index=False )
df12.to_excel(writer3,backlog12, index=False )
df13.to_excel(writer3,backlog13, index=False )
df14.to_excel(writer3,backlog14, index=False )
df15.to_excel(writer3,backlog15, index=False )
df16.to_excel(writer3,backlog16, index=False )
df17.to_excel(writer3,backlog17, index=False )
df18.to_excel(writer3,backlog18, index=False )
df19.to_excel(writer3,backlog19, index=False )
df20.to_excel(writer3,backlog20, index=False )
#==================================================================
workbook7=writer3.book
worksheet7=writer3.sheets[backlog7]
(max_row, max_col) = df7.shape

# Create a list of column headers, to use in add_table().
column_settings7 = []
for header in df7.columns:
    column_settings7.append({'header': header})

# Add the table.
worksheet7.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings7})

# Make the columns wider for clarity.
worksheet7.set_column(0, max_col - 1, 12)
#==================================================================
workbook8=writer3.book
worksheet8=writer3.sheets[backlog8]
(max_row, max_col) = df8.shape

# Create a list of column headers, to use in add_table().
column_settings8 = []
for header in df8.columns:
    column_settings8.append({'header': header})

# Add the table.
worksheet8.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings8})

# Make the columns wider for clarity.
worksheet8.set_column(0, max_col - 1, 12)
#==================================================================
workbook9=writer3.book
worksheet9=writer3.sheets[backlog9]
(max_row, max_col) = df9.shape

# Create a list of column headers, to use in add_table().
column_settings9 = []
for header in df9.columns:
    column_settings9.append({'header': header})

# Add the table.
worksheet9.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings9})

# Make the columns wider for clarity.
worksheet9.set_column(0, max_col - 1, 12)
#==================================================================
workbook10=writer3.book
worksheet10=writer3.sheets[backlog10]
(max_row, max_col) = df3.shape

# Create a list of column headers, to use in add_table().
column_settings10 = []
for header in df10.columns:
    column_settings10.append({'header': header})

# Add the table.
worksheet10.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings10})

# Make the columns wider for clarity.
worksheet10.set_column(0, max_col - 1, 12)
#==================================================================
workbook11=writer3.book
worksheet11=writer3.sheets[backlog11]
(max_row, max_col) = df11.shape

# Create a list of column headers, to use in add_table().
column_settings11 = []
for header in df11.columns:
    column_settings11.append({'header': header})

# Add the table.
worksheet11.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings11})

# Make the columns wider for clarity.
worksheet11.set_column(0, max_col - 1, 12)
#==================================================================
workbook12=writer3.book
worksheet12=writer3.sheets[backlog12]
(max_row, max_col) = df12.shape

# Create a list of column headers, to use in add_table().
column_settings12 = []
for header in df12.columns:
    column_settings12.append({'header': header})

# Add the table.
worksheet12.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings12})

# Make the columns wider for clarity.
worksheet12.set_column(0, max_col - 1, 12)
#==================================================================
workbook13=writer3.book
worksheet13=writer3.sheets[backlog13]
(max_row, max_col) = df13.shape

# Create a list of column headers, to use in add_table().
column_settings13 = []
for header in df13.columns:
    column_settings13.append({'header': header})

# Add the table.
worksheet13.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings13})

# Make the columns wider for clarity.
worksheet13.set_column(0, max_col - 1, 12)
#==================================================================
workbook14=writer3.book
worksheet14=writer3.sheets[backlog14]
(max_row, max_col) = df14.shape

# Create a list of column headers, to use in add_table().
column_settings14 = []
for header in df14.columns:
    column_settings14.append({'header': header})

# Add the table.
worksheet14.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings14})

# Make the columns wider for clarity.
worksheet14.set_column(0, max_col - 1, 12)
#==================================================================
workbook15=writer3.book
worksheet15=writer3.sheets[backlog15]
(max_row, max_col) = df15.shape

# Create a list of column headers, to use in add_table().
column_settings15 = []
for header in df15.columns:
    column_settings15.append({'header': header})

# Add the table.
worksheet15.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings15})

# Make the columns wider for clarity.
worksheet15.set_column(0, max_col - 1, 12)
#==================================================================
workbook16=writer3.book
worksheet16=writer3.sheets[backlog16]
(max_row, max_col) = df16.shape

# Create a list of column headers, to use in add_table().
column_settings16 = []
for header in df16.columns:
    column_settings16.append({'header': header})

# Add the table.
worksheet16.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings16})

# Make the columns wider for clarity.
worksheet16.set_column(0, max_col - 1, 12)
#==================================================================
workbook17=writer3.book
worksheet17=writer3.sheets[backlog17]
(max_row, max_col) = df17.shape

# Create a list of column headers, to use in add_table().
column_settings17 = []
for header in df17.columns:
    column_settings17.append({'header': header})

# Add the table.
worksheet17.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings17})

# Make the columns wider for clarity.
worksheet17.set_column(0, max_col - 1, 12)
#==================================================================
workbook18=writer3.book
worksheet18=writer3.sheets[backlog18]
(max_row, max_col) = df18.shape

# Create a list of column headers, to use in add_table().
column_settings18 = []
for header in df18.columns:
    column_settings18.append({'header': header})

# Add the table.
worksheet18.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings18})

# Make the columns wider for clarity.
worksheet18.set_column(0, max_col - 1, 12)
#==================================================================
workbook19=writer3.book
worksheet19=writer3.sheets[backlog19]
(max_row, max_col) = df19.shape

# Create a list of column headers, to use in add_table().
column_settings19 = []
for header in df19.columns:
    column_settings19.append({'header': header})

# Add the table.
worksheet19.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings19})

# Make the columns wider for clarity.
worksheet19.set_column(0, max_col - 1, 12)
#==================================================================
workbook20=writer3.book
worksheet20=writer3.sheets[backlog20]
(max_row, max_col) = df20.shape

# Create a list of column headers, to use in add_table().
column_settings20 = []
for header in df20.columns:
    column_settings20.append({'header': header})

# Add the table.
worksheet20.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings20})

# Make the columns wider for clarity.
worksheet20.set_column(0, max_col - 1, 12)

writer3.save()
writer3.close()