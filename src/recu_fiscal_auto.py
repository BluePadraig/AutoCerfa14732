import datetime
import importlib.metadata
import math
import os
import pathlib
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

from fillpdf import fillpdfs
import pandas as pd


# todo: Récupération du package name dans le pyproject.toml
project_package_name = "recu_fiscal_auto"

# todo: Récupération de la version dans le pyproject.toml
default_version = "1.0.0"
try:
    version = importlib.metadata.version(project_package_name)
except importlib.metadata.PackageNotFoundError as e:
    version = default_version

project_url = "https://github.com/BluePadraig/RecuFiscalAuto/blob/main/README.md"

ok_to_continue = messagebox.askokcancel(
    title=f"Bienvenue dans le générateur de reçus V{version}",
    message=f"La documentation est disponible là : {project_url} \n Voulez-vous continuer ?"
)
if not ok_to_continue:
    exit()

main_window = tk.Tk()
main_window.withdraw()  # Cacher la fenêtre principale

input_pdf_path = filedialog.askopenfilename(
    title="Sélectionnez le modèle de reçu pdf qui contient les champs à remplir",
    filetypes=[("Fichiers PDF", "*.pdf")]
)
if not input_pdf_path:
    exit()
form_fields = fillpdfs.get_form_fields(input_pdf_path=input_pdf_path, sort=False, page_number=None)

output_path = filedialog.askdirectory(title="Dossier où créer les reçus pdf remplis")
if not output_path:
    exit()
path_to_create = pathlib.Path(output_path)
if not path_to_create.exists():
    path_to_create.mkdir()

output_pdf_description = simpledialog.askstring(
    title="Demande de description",
    prompt="Description à ajouter dans le nom des pdf remplis :",
    initialvalue="Reçu fiscal",
)
if not output_pdf_description:
    output_pdf_description = ""

input_path = pathlib.Path(input_pdf_path).parent
input_xlsx = filedialog.askopenfilename(
    title="Sélectionnez le fichier Microsoft Excel ou LibreOffice Calc contenant une ligne de données par pdf à remplir",
    initialdir=input_path,
    filetypes=[("Fichiers Excel", "*.xlsx"), ("Fichiers Calc", "*.ods")]
)
if not input_xlsx:
    exit()

nom_colonnes = [
    "Numéro d'ordre du reçu",
    "Nom",
    "Prénom",
    "Adresse",
    "Code postal",
    "Commune",
    "Montant du don (chiffres)",
    "Date du don",
    "Nature",
]
df = pd.read_excel(input_xlsx, sheet_name="Donateurs", skiprows=1, usecols=nom_colonnes, dtype={'Code postal': str})
df = df.fillna("")

date_recu_for_path = str(datetime.date.today())
yy, mm, dd = date_recu_for_path.split("-")
date_recu_francais = f"{dd}/{mm}/{yy}"

correspondance_recu_pdf = {
    "Numéro d'ordre du reçu": "Numero ordre du recu",
    "Nom": "Nom",
    "Prénom": "Prenom",
    "Adresse": "Adresse",
    "Code postal": "Code postal",
    "Commune": "Commune",
    "Montant du don (chiffres)": "Montant du don",
    "Date du don":  "Date du don",
    "Nature": "Nature",
    "Date du reçu": "Date du recu",
}


def convert_df_to_pdf(input_df: pd.Series) -> dict:
    output_dict = {}
    for key, value in input_df.items():
        if key in correspondance_recu_pdf:
            output_dict[correspondance_recu_pdf[key]] = value
    return output_dict

def clean_code_postal(code_postal: float|str) -> str:
    if type(code_postal) is float:
        if math.isnan(code_postal):
            return ""
        return str(int(code_postal))
    return code_postal

def clean_date_du_don(time_stamp: pd.Timestamp) -> str:
    return time_stamp.strftime("%d/%m/%Y")

def clean_montant_du_don(montant: float) -> str:
    return str(montant).replace(".", ",")

nb_recus = 0
for index, don in df.iterrows():
    num_recu = don["Numéro d'ordre du reçu"]
    nom = don["Nom"]
    prenom = don["Prénom"]

    output_pdf = (f"{output_path}/{num_recu}-{output_pdf_description}-"
                  f"{prenom}-{nom}-{date_recu_for_path}.pdf")

    dict_pdf = convert_df_to_pdf(don)
    dict_pdf["Date du recu"] = date_recu_francais
    dict_pdf["Code postal"] = clean_code_postal(dict_pdf["Code postal"])
    dict_pdf["Date du don"] = clean_date_du_don(dict_pdf["Date du don"])
    dict_pdf["Montant du don"] = clean_montant_du_don(dict_pdf["Montant du don"])
    make_the_pdf_uneditable  = True
    fillpdfs.write_fillable_pdf(
        input_pdf_path, output_pdf, dict_pdf, flatten=make_the_pdf_uneditable)
    nb_recus += 1


messagebox.showinfo(
    title="Génération des reçus terminée",
    message=f"{nb_recus} reçus générés dans {output_path}"
)

os.startfile(output_path)
