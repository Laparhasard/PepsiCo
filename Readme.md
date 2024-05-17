
# Documentation pour l'utilisation du programme d'analyse d'inventaire

## Introduction
Ce programme d'analyse d'inventaire permet de traiter des fichiers Excel, d'agréger les données et de générer un graphique des références les plus importantes par quantité d'inventaire (`INV_Q_COLIS`). Le programme utilise une interface graphique pour faciliter l'entrée des paramètres par l'utilisateur.

## Prérequis

### Installation de Python
1. **Windows**
    - Téléchargez Python depuis le site officiel : [python.org](https://www.python.org/downloads/)
    - Exécutez l'installateur et assurez-vous de cocher l'option "Ajouter Python à PATH" lors de l'installation.

2. **Linux**
    - Utilisez le gestionnaire de paquets de votre distribution pour installer Python. Par exemple, sur Ubuntu :
      ```bash
      sudo apt update
      sudo apt install python3 python3-pip
      ```

3. **Mac**
    - Utilisez Homebrew pour installer Python :
      ```bash
      /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
      brew install python
      ```

### Installation des dépendances
Vous devez installer les bibliothèques Python suivantes : `pandas`, `matplotlib`, et `tkinter`.

1. **Windows, Linux, Mac**
    - Utilisez `pip` pour installer les bibliothèques :
      ```bash
      pip install pandas matplotlib tk
      ```

## Préparation des fichiers
Avant d'exécuter le programme, vous devez vous assurer que les chemins de fichier (`file_path`) et le répertoire (`directory`) aux lignes 93 et 130 pointent vers les bons fichiers sur votre système.

### Modification des chemins de fichier
1. Ouvrez le fichier Python dans un éditeur de texte.
2. Modifiez les lignes suivantes pour qu'elles correspondent aux emplacements de vos fichiers :

    ```python
    directory = '/chemin/vers/votre/répertoire/Stock Inventaire OK/'
    sales_file_path = '/chemin/vers/votre/fichier/Ventes 2023-2024.xlsx'
    ```

### Exemple de modification
Si vos fichiers sont situés dans le dossier `C:\Documents\Inventaire\`, vous devez changer les lignes comme suit :

```python
directory = 'C:/Documents/Inventaire/Stock Inventaire OK/'
sales_file_path = 'C:/Documents/Inventaire/Ventes 2023-2024.xlsx'
```

### Lignes à modifier dans le code
Voici un extrait du code montrant où effectuer ces modifications :

```python
# Répertoire contenant les fichiers Excel
directory = '/chemin/vers/votre/répertoire/Stock Inventaire OK/'

# Chemin du fichier de ventes
sales_file_path = '/chemin/vers/votre/fichier/Ventes 2023-2024.xlsx'
```

## Exécution du programme
Pour exécuter le programme, suivez les étapes ci-dessous :

1. **Lancer le programme**
    - Ouvrez une console (cmd, PowerShell, terminal) et exécutez le script Python :
      ```bash
      python path_to_your_script.py
      ```
    - Remplacez `path_to_your_script.py` par le chemin vers votre fichier Python.

2. **Utilisation de l'interface graphique**
    - Une fenêtre s'ouvrira avec deux champs de saisie et un bouton "Soumettre".
    - **Champ "Année"** : Entrez l'année des données que vous souhaitez analyser (par exemple, 2023 ou 2024).
    - **Champ "Interval Semaines"** : Entrez l'intervalle de semaines à analyser au format `SXX-SXX` (par exemple, `1-17` pour analyser les semaines de 1 à 17).
    - Cliquez sur le bouton "Soumettre" pour lancer l'analyse.

### Résultats
- **Données agrégées** : Les données agrégées seront enregistrées dans un fichier CSV situé dans le même répertoire que vos fichiers d'inventaire.
- **Graphique** : Un graphique en barres empilées sera généré et enregistré dans un fichier image PNG dans le même répertoire.

## Points supplémentaires
- **Message d'erreur et d'avertissement** : Le programme affichera des messages d'erreur ou d'avertissement si des fichiers sont manquants ou si des colonnes nécessaires ne sont pas trouvées.
- **Ajustement des chemins de fichier** : Assurez-vous que les chemins des fichiers et des répertoires sont corrects avant d'exécuter le programme.

## Exemple de configuration
Si vous utilisez Windows et que vos fichiers sont situés dans `C:\Documents\Inventaire\`, votre script pourrait ressembler à ceci :

```python
directory = 'C:/Documents/Inventaire/Stock Inventaire OK/'
sales_file_path = 'C:/Documents/Inventaire/Ventes 2023-2024.xlsx'
```

Lancez le programme avec la commande suivante dans une console :

```bash
python C:/Documents/Inventaire/analyse_inventaire.py
```

Entrez ensuite par exemple `2024` pour l'année et `1-17` pour l'intervalle de semaines dans l'interface graphique pour prendre en compte de la semaine 1 à 17 (il faudra attendre un peu que le programme parcours toutes les données).

## Conclusion
Ce programme facilite l'analyse d'inventaire en agrégeant les données de plusieurs fichiers Excel et en générant des graphiques pour visualiser les résultats. Suivez les instructions ci-dessus pour installer les dépendances, configurer les chemins de fichier, et exécuter le programme. Pour toute question ou problème, n'hésitez pas à consulter la documentation officielle de Python et des bibliothèques utilisées.
```