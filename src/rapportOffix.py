""" 
    @File : rapportOffix.py
    @Brief : Analyse du rapport offix
    @Author : Jonathan Dagnault et Ali Boustta
    @Version : 1.0
    @Date : 19-06-23
"""
"""
Hiérarchie du fichier XML

root
|<Details Level="1">                        [0]
||<Section SectionNumber="0">               [0]
|||<Subreport Name="SommaireFacture">       [0]
||||<Group Level="1">                       *** C'est à ce niveau que chaque facture est divisée
|||||<Details Level="2">                    [1]
||||||<Section SectionNumber="0">           [0]
|||||||                                     *** C'est à ce niveau que la data de chaque facture est divisée
"""

import xml.etree.ElementTree as ET
import path

#Tags XML 
GROUP = "{urn:crystal-reports:schemas:report-detail}Group"


"""
    @Brief : Retourne le subreport "Sommaire des factures" du fichier xml qui provient de Offix
    @Return : un subreport ('{urn:crystal-reports:schemas:report-detail}Subreport') contenant les données sur les facturations
"""
def getOffixSubreportFactures():

    INPUT_FILE = "rapportOffix.xml"

    #Ouvre et parse le fichier  qui vient de Offix
    tree = ET.parse(path.getDataPath(INPUT_FILE))
    root = tree.getroot()

    #Retourne le subreport qui contient la data nécessaire
    return root[0][0][0]


"""
    @Brief : Itère à travers le subreport pour trouver la data sur chaque facturation (numéro, prix, client, poste)
    @Return : Une liste de toutes les facturations (numéro, prix, client, poste)
"""
def getOffixFacturesData():
    
    subreport = getOffixSubreportFactures()

    OffixFacturesData = []

    #On itère à travers toutes les factures offix
    for facture in subreport.iter(GROUP):

        #On initialise la facture
        facturation = {'Numéro Facture Offix' : None,
                        'Prix Offix' : None,
                        'Nom Client Offix' : None,
                        'Poste Offix' : None
        }

        #On itère à travers les données de chaque facture
        for value_index, value in enumerate(facture[1][0].iter()):
            
            #La 2e valeur de la facture représente le numéro de facture
            if value_index == 2: 
                facturation['Numéro Facture Offix'] = value.text
            
            #La 11e valeur de la facture représente le prix de la facture
            elif value_index == 11:
                facturation['Prix Offix'] = value.text
            
            #La 14e valeur de la facture représente le nom du client
            elif value_index == 14:
                facturation['Nom Client Offix'] = value.text
            
            #La 17e valeur de la facture représente le poste
            elif value_index == 17:
                facturation['Poste Offix'] = value.text
        
        OffixFacturesData.append(facturation)

    return OffixFacturesData


"""
     @Brief : Insère les données dans le fichier de comparaison
"""
def insertOffixData(sheet, offix_data):

    #Spécifie les colonnes où l'on veut insérer les données
    column_mapping = {
        'A': 'Numéro Facture Offix',
        'B': 'Prix Offix',
        'C': 'Nom Client Offix',
        'D': 'Poste Offix'
    }

    #On commence à la 2e ligne
    row = 2

    #Itère à travers les commandes
    for facture in offix_data:

        #Itère les quatres colonnes où les données doivent être insérées
        for column, column_header in column_mapping.items():
            
            #Remplace la valeur de la cell par celle de la facture
            sheet.range(f"{column}{row}").value = facture[column_header]  

        #Change de ligne lorsque la facture est insérée
        row += 1