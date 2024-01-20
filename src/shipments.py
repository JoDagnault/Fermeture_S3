""" 
    @File : shipments.py
    @Brief : Analyse du fichier 'shipments.xml' qui provient de magento
    @Author : Jonathan Dagnault et Ali Boustta
    @Version : 1.0
    @Date : 18-06-23
"""

import xml.etree.ElementTree as ET
import xlwings as xw
import path

#Tags XML 
ROW = "{urn:schemas-microsoft-com:office:spreadsheet}Row"
CELL = "{urn:schemas-microsoft-com:office:spreadsheet}Cell"
DATA = "{urn:schemas-microsoft-com:office:spreadsheet}Data"


"""
    @Brief : Trouve la table excel du fichier "shipments.xml" qui provient de Magento
    @Return : une table excel ('{urn:schemas-microsoft-com:office:spreadsheet}Table') contenant les données sur les expéditions
"""
def getShipmentsTable():
    
    INPUT_FILE = "shipments.xml"

    #Ouvre et parse le fichier xml passé en paramètre
    tree = ET.parse(path.getDataPath(INPUT_FILE))
    root = tree.getroot()

    #Retourne la table XML qui contient la data nécessaire
    return root[2][0]


"""
    @Brief : Itère à travers la table pour trouver la data sur chaque commande (numéro, client, prix)
    @Return : Une liste de toutes les commandes (numéro, client, prix) expédiées 
"""
def getShipmentsData():
    
    table = getShipmentsTable()

    shipmentsData = []

    #On itere a travers les rangées de la table (qui représentent les commandes expédiées)
    for row_index, row in enumerate(table.iter(ROW)):
    
        #On saute la première ligne qui contient les titres de la table
        if row_index != 0 :
            
            #Initialise la commande
            commande = {'Numéro Commande Magento' : None,
                        'Nom Client Magento' : None,
                        'Prix Magento' : None
            }   

            for cell_index, cell in enumerate(row.iter(CELL)):

                #La 2e cell représente le numéro de commande
                if cell_index == 2 :
                    commande['Numéro Commande Magento'] = cell.find(DATA).text
                
                #La 4e cell représente le nom du client
                elif cell_index == 4 :
                    commande['Nom Client Magento'] = cell.find(DATA).text

                #La 6e cell représente le prix de la commande, qui doit être formatté à deux décimales
                elif cell_index == 6 :
                    commande['Prix Magento'] = cell.find(DATA).text[:-2]

            shipmentsData.append(commande)

    return shipmentsData


"""
     @Brief : Insère les données dans le fichier de comparaison
"""
def insertShipmentsData(sheet, shipments_data):

    #Spécifie les colonnes où l'on veut insérer les données
    column_mapping = {
        'E': 'Numéro Commande Magento',
        'F': 'Nom Client Magento',
        'G': 'Prix Magento'
    }

    #On commence à la 2e ligne
    row = 2

    #Itère à travers les commandes
    for commande in shipments_data:

        #Itère les trois colonnes où les données doivent être insérées
        for column, column_header in column_mapping.items():
            
            #Remplace la valeur de la cell par celle de la commande
            sheet.range(f"{column}{row}").value = commande[column_header]  

        #Change de ligne lorsque la commande est insérée
        row += 1