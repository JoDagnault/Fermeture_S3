""" 
    @File : path.py
    @Brief : Fichier contenant des fonctions de base servant à la recherche de path
    @Author : Jonathan Dagnault et Ali Boustta
    @Version : 1.0
    @Date : 05-07-23
"""

import os.path as path

"""
    @Brief : Trouve le dossier parent (directory) du fichier spécifié en paramètre
    @Param[in] file_path : Path absolue du fichier dont on veut le dossier parent (directory)
    @Return : Path du dossier parent (directory) du fichier spécifié en paramètre
"""
def getDirectory(file_path):
    return path.dirname(file_path)


"""
    @Brief : Trouve le dossier parent (directory) du projet complet
    @Return : La path absolue du dossier parent (directory) du projet
"""
def getProjetDirectory():
    #Chaque fichier source est situé dans le dossier /src qui est lui situé dans le dossier parent
    #Il faut donc trouver le dossier parent (directory) au dossier parent (directory) du fichier
    return getDirectory(getDirectory(__file__))


"""
    @Brief : Donne la path absolue d'un fichier dans un dossier parent (directory) spécifié
    @Param[in] dir : Dossier parent (directory) où se trouve le fichier dont on veut la path absolue
    @Param[in] rel_path : Path relative du fichier spécifié par rapport au dossier parent (directory)
    @Return : La path absolue du fichier spécifié
"""
def getAbsolutePath(dir, rel_path):
    return path.join(dir, rel_path)


"""
    @Brief : Trouve la path absolue d'un fichier contenu dans le dossier data
    @Param[in] file_name : Nom du fichier dont on veut la path absolue
    @Return : La path absolue du fichier spécifié
"""
def getDataPath(file_name):
    projet_dir = getProjetDirectory()
    data_dir = getAbsolutePath(projet_dir, "data/")
    return getAbsolutePath(data_dir, f"{file_name}")


"""
    @Brief : Trouve la path absolue d'un fichier contenu dans le dossier output
    @Param[in] file_name : Nom du fichier dont on veut la path absolue
    @Return : La path absolue du fichier spécifié
"""
def getOutputPath(file_name):
    projet_dir = getProjetDirectory()
    output_dir = getAbsolutePath(projet_dir, "output/")
    return getAbsolutePath(output_dir, f"{file_name}")


"""
    @Brief : Trouve la path absolue d'un fichier contenu dans le dossier templates
    @Param[in] file_name : Nom du fichier dont on veut la path absolue
    @Return : La path absolue du fichier spécifié
"""
def getTemplatePath(file_name):
    projet_dir = getProjetDirectory()
    templates_dir = getAbsolutePath(projet_dir, "templates/")
    return getAbsolutePath(templates_dir, f"{file_name}")