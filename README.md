# RecuFiscalAuto
Remplissage automatique de reçus fiscaux en pdf à partir d'un fichier xlsx ou ods et d'un modèle en pdf

## Mode d'emploi

* Télécharger le programme recu_fiscal_auto.exe sur [GitHub](https://github.com/BluePadraig/RecuFiscalAuto/releases/)
  * Dans la rubrique Releases Sélectionner la version la plus récente
  * Sélectionner Assets
  * Télécharger le fichier recu_fiscal_auto.exe
* Exécuter le programme recu_fiscal_auto.exe
  * Le programme prend quelques secondes avant d'afficher la première boite de dialogue 
* Indiquer le chemin du modèle de reçu fiscal, au format pdf
* Indiquer le chemin du dossier où le programme doit générer les reçus
  * Tout reçu déjà existant dans ce dossier sera écrasé si un nouveau reçu est généré avec les mêmes informations donateur, notamment avec le même numéro d'ordre
* Indiquer la description à ajouter dans le nom des pdf remplis
  * Ce texte doit être court, pour éviter les noms de fichier à rallonge
  * Ce texte ne doit pas contenir de caractère créant un conflit dans un nom de fichier
    * Par exemple, pas de \ ou / ou ?
* Indiquer le chemin du fichier les informations des donateurs
* Voir le [format des informations donateurs](#format-des-informations-donateurs)
* Le programme va générer un reçu fiscal par donateur, au format pdf
  * Le nom des fichiers générés est au format suivant
    * Numéro d'ordre du reçu-Description-Prénom-Nom-Date de génération du reçu.pdf
  * Exemple
    * 2409230-Reçu fiscal Les Amis du Libre-Harry-COVERT-2024-09-29.pdf

## Code source

Le programme est écrit en Python  
Le code source est disponible librement sur [GitHub](https://github.com/BluePadraig/RecuFiscalAuto)  
Il est soumis à la licence GNU GPL
Ce programme est utilisable gratuitement et librement  
Vous êtes libre de modifier ce programme, tant qu'il conserve sa licence libre

Les contributions sous forme d'amélioration de code sont les bienvenues

L'exécutable fourni est conçu pour Windows.
Mais il est tout à fait possible d'exécuter le programme directement avec un interpréteur Python, sous Windows, Linux ou MacOs

## Construire le fichier exécutable

En cas de modification du code source, vous pouvez reconstruire le fichier exécutable : 
```cmd
pyinstaller --onefile .src/recu_fiscal_auto.py
```

## Format du modèle de reçu fiscal

Ce modèle doit-être un fichier pdf, avec les champs éditables suivants :

* "Numero ordre du recu"
* "Nom"
* "Prenom"
* "Adresse"
* "Code postal" (de type nombre)
* "Commune"
* "Montant du don" (de type nombre)
* "Date du don" (de type date)
* "Nature"
* "Date du recu" (de type date)

Pour éviter les problèmes d'encodage, les noms des champs n'ont pas d'accent.

Pour créer ce modèle de reçu fiscal, vous pouvez utiliser le logiciel Writer de la suite LibreOffice.
Vous pouvez vous inspirer du fichier fourni en exemple : [Exemple de modèle de reçu fiscal.odt](</Exemple%20de%20modèle%20de%20reçu%20fiscal.odt>)

# Format des informations donateurs

Les informations sur les donateurs doivent être renseignées dans un fichier de type Microsoft Excel (.xlsx) ou LibreOffice Calc (.ods).

### Contraintes à respecter 

* Une ligne par donateur
* Pas de ligne vide
* Les informations des donateurs sont dans un onglet (feuille) nommé Donateurs
* Les colonnes ci-dessous doivent être présentes, en respectant très exactement leur nom
  * Numéro d'ordre du reçu
  * Nom	
  * Prénom	
  * Adresse	
  * Code postal	
  * Commune
  * Montant du don (chiffres)
  * Date du don	
  * Nature
* Les numéros d'ordre de reçu doivent être uniques. C'est ce qui garanti qu'il n'y a pas de confusion entre les reçus. 
* Il est possible d'ajouter d'autres colonnes. Elles seront simplement ignorées par le programme de génération de reçus