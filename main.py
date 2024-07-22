# import pandas as pd
# import pdfkit
# from jinja2 import Environment, FileSystemLoader
# import os
# import imgkit

# # Lire le fichier Excel
# df = pd.read_excel("./Evaluation.xlsx")

# # Liste des colonnes à exclure
# excluded_columns = ["Nom_Audio", "Qualification de l'appel", "Score de l'appel", "Points positifs", "Points négatifs", "Résumé de l'agent"]


# def create_evaluation_pdf(output_file: str, line_from_df: pd.Series):
    
#     options = {
#     'page-size': 'A4',
#     'orientation': 'Landscape',
#     'margin-top': '10mm',
#     'margin-right': '10mm',
#     'margin-bottom': '10mm',
#     'margin-left': '10mm',
#     'encoding': 'UTF-8',
#     'no-outline': None,
#     'enable-local-file-access': None,
#     }
#     # Créer un environnement Jinja2 pour le modèle HTML
#     env = Environment(loader=FileSystemLoader('.'))
#     template = env.get_template('template.html')

#     # Filtrer les colonnes valides
#     valid_criteria = [(column, line_from_df[column]) for column in df.columns if column not in excluded_columns]

#     # Créer le contexte pour le modèle HTML
#     context = {
#         'image_path': "./logo.png",
#         'agent_name': "Yassir El Mansouri",
#         'offre_proposee': "Pack",
#         'resultat_vente': "ko",
#         'valid_criteria': valid_criteria,
#         'points_positifs': line_from_df['Points positifs'].split('- '),
#         'points_negatifs': line_from_df['Points négatifs'].split('- '),
#         'resume_agent': line_from_df["Résumé de l'agent"]
#     }

#     # Rendre le modèle avec le contexte
#     html_out = template.render(context)

#     # Sauvegarder le HTML temporaire
#     with open('temp.html', 'w' , encoding='utf-8') as f:
#         f.write(html_out)

#     # Chemin vers wkhtmltopdf.exe
#     path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

#     # Options pdfkit pour spécifier le chemin de wkhtmltopdf
#     config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
#     config_img = imgkit.config("C:\Program Files\wkhtmltopdf\bin")
#     # Convertissez l'HTML en image
#     imgkit.from_file('temp.html', 'evaluation.png', config=config_img, options={'width': 1500})

#     # Convertissez l'image en PDF
#     imgkit.from_file('evaluation.png', output_file , configuration=config, options=options)
    
#     # Convertir le HTML en PDF
#     # pdfkit.from_file('temp.html', output_file, configuration=config , options=options )

# # Utilisation de la fonction
# create_evaluation_pdf("agent.pdf", df.loc[0])

import pandas as pd
import pdfkit
from jinja2 import Environment, FileSystemLoader
import imgkit
import os

# Lire le fichier Excel
df = pd.read_excel("./Evaluation.xlsx")

# Liste des colonnes à exclure
excluded_columns = ["Nom_Audio", "Qualification de l'appel", "Score de l'appel", "Points positifs", "Points négatifs", "Résumé de l'agent"]

def create_evaluation_pdf(output_file: str, line_from_df: pd.Series):
    
    options = {
        'page-size': 'A4',
        'orientation': 'Landscape',
        'margin-top': '10mm',
        'margin-right': '10mm',
        'margin-bottom': '10mm',
        'margin-left': '10mm',
        'encoding': 'UTF-8',
        'no-outline': None,
        'enable-local-file-access': None,
    }
    op = { 
        'width': 1500  ,  
        'enable-local-file-access': ''  # Enable local file access 
    }
    # Créer un environnement Jinja2 pour le modèle HTML
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')

    # Filtrer les colonnes valides
    valid_criteria = [(column, line_from_df[column]) for column in df.columns if column not in excluded_columns]

    # Créer le contexte pour le modèle HTML
    context = {
        'image_path': "./logo.png",
        'agent_name': "Yassir El Mansouri",
        'offre_proposee': "Pack",
        'resultat_vente': "ko",
        'valid_criteria': valid_criteria,
        'points_positifs': line_from_df['Points positifs'].split('- '),
        'points_negatifs': line_from_df['Points négatifs'].split('- '),
        'resume_agent': line_from_df["Résumé de l'agent"]
    }

    # Rendre le modèle avec le contexte
    html_out = template.render(context)

    # Sauvegarder le HTML temporaire
    with open('temp.html', 'w', encoding='utf-8') as f:
        f.write(html_out)

    # Chemins vers les exécutables
    path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    path_to_wkhtmltoimage = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe'

    # Configurations pour pdfkit et imgkit
    pdfkit_config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
    imgkit_config = imgkit.config(wkhtmltoimage=path_to_wkhtmltoimage)

    # Convertir le HTML en PDF directement
    # pdfkit.from_file('temp.html', output_file, configuration=pdfkit_config, options=options)
    
    # Si vous avez besoin de convertir HTML en image (en cas de besoin spécifique), utilisez :
    imgkit.from_file('temp.html', 'Evaluation.png', config=imgkit_config, options=op)
    # Ensuite, si vous souhaitez convertir l'image en PDF, utilisez :
    imgkit.from_file('Evaluation.png', output_file, config=imgkit_config)

# Utilisation de la fonction
create_evaluation_pdf("agent.pdf", df.loc[0])
