""" 
    @File : fermeture_S3.py
    @Brief : Combine les données de Magento et de Offix et insère dans le fichier de comparaison
    @Author : Jonathan Dagnault et Ali Boustta
    @Version : 1.0
    @Date : 02-07-23
"""

import xlwings as xw
import path
import shipments
import rapportOffix


"""
    @Brief : Prend les données fournies et les combinent dans le fichier de comparaison
"""
def dataToExcel(template_file_name, output_file_name, sheet_name, shipments_data, offix_data):

    #Charge le workbook
    workbook = xw.Book(path.getTemplatePath(template_file_name))

    #Sélectionne la sheet
    sheet = workbook.sheets[sheet_name]

    #Insère les données dans le fichier
    shipments.insertShipmentsData(sheet, shipments_data)
    rapportOffix.insertOffixData(sheet, offix_data)

    #Sauvegarde le workbook à la fin
    workbook.save(path.getOutputPath(output_file_name))

    #Refresh la data pour update les tables power query
    workbook.api.RefreshAll()

    #workbook.close()


"""
-------------------------------------------------------------------------------
    Utilisation
-------------------------------------------------------------------------------
"""
def main():
    TEMPLATE_FILE = "Output_Fermeture_S3_Template.xlsx"
    OUTPUT_FILE = "Output_Fermeture_S3.xlsx"
    SHEET_NAME = "Ensemble des tableaux"

    shipments_data = shipments.getShipmentsData()
    offix_data = rapportOffix.getOffixFacturesData()

    dataToExcel(TEMPLATE_FILE, OUTPUT_FILE, SHEET_NAME, shipments_data, offix_data)

if __name__ == '__main__':
    main()