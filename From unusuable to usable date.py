#The idea of this algorithm is to extract usefull #information from a Excel with unsuable form to analyze and #transform in a usuable way. Adding a graph to have a fast
#view of price evolution.

import pandas as pd
import openpyxl
import numpy as np
import xlsxwriter
import matplotlib.pyplot as plt

#getting the file to analyze
file_base = pd.read_excel("azucar actualizado.xlsx")
#convert in dataframe
file_df = pd.DataFrame(file_base)

#getting data I want from specific location on excel
fechas1 = file_df.iloc[688:719, 0]
datos1 = file_df.iloc[688:719, 3]

fechas2 = file_df.iloc[688:719, 6]
datos2 = file_df.iloc[688:719, 9]

fechas3 = file_df.iloc[688:719, 12]
datos3 = file_df.iloc[688:719, 15]

#Make a new dataframe with all the usefull dates
fecha = pd.DataFrame(
  fechas1.append(fechas2, ignore_index=True).append(fechas3,
                                                    ignore_index=True))

#transform date format
fecha[0] = fecha[0].dt.strftime('%d/%m/%y')

#make a new DF with all values required
precio = pd.DataFrame(
  datos1.append(datos2, ignore_index=True).append(datos3, ignore_index=True))

#join dates & values through "concat"
nuevo_df = pd.concat({'Fecha': fecha, 'Precio': precio}, axis=1)

# Deleting NaN values
nuevo_df = nuevo_df.dropna()

#Creation of excel file where I'm going to save new information
file_sugar = pd.ExcelWriter("azucartucumanconsolidado.xlsx",
                            engine='xlsxwriter')

nuevo_df.to_excel(file_sugar,
                  sheet_name="general",
                  index=True,
                  engine='xlsxwriter')

#Creation of line chart to have a view of price evolution.
workbook = file_sugar.book
worksheet = file_sugar.sheets['general']

chart = workbook.add_chart({'type': 'line'})
chart.add_series({
  'name': 'Azucar Tucuman',
  'categories': 'general!B4:B50',
  'values': 'general!C4:C50'
})

# Configure the chart axes.
chart.set_x_axis({'name': 'Fecha', 'position_axis': 'on_tick'})
chart.set_y_axis({
  'name': 'Precio [ARS/KG]',
  'major_gridlines': {
    'visible': False
  }
})

# Set chart dimensions
chart.set_size({'width': 720, 'height': 476})
# Turn off chart legend. It is on by default in Excel.
chart.set_legend({'position': 'none'})

# Insert chart into worksheet
worksheet.insert_chart('F5', chart)

#Save the file
file_sugar.save()
