import pandas as pd
import os
import matplotlib.pyplot as plt
from tkinter import Tk, Label, Entry, Button, messagebox

# Fonction pour extraire la date du nom de fichier et la formater en Semaine.Année
def extract_date_from_filename(filename):
    basename = os.path.basename(filename)
    parts = basename.split()
    year = parts[3]
    week = parts[4][1:].replace('.xlsx', '')  # Supprimer l'extension ".xlsx"
    return f"{week}.{year}"

# Fonction pour traiter un fichier unique et retourner les données
def process_file(file_path):
    try:
        # Charger les données du fichier fourni, en assurant des noms de colonnes cohérents
        df = pd.read_excel(file_path, header=3)  # Utiliser header=3 pour lire les en-têtes à partir de la 4ème ligne
        df.columns = df.columns.str.strip()  # Supprimer les espaces en début et fin de chaîne dans les noms de colonnes

        # Mappages des noms de colonnes pour corriger les variations et les erreurs typographiques
        column_mappings = {
            'Ref & Description': 'REF & DESCRIPTION',
            'REF & DESRCIPTION': 'REF & DESCRIPTION',
            'inv q colis': 'INV_Q_COLIS',
            'Inventory Q PAL': 'INV_Q_COLIS'
        }
        df.rename(columns=column_mappings, inplace=True)

        # Vérifier la présence des colonnes nécessaires avant de continuer
        necessary_columns = ['REF & DESCRIPTION', 'INV_Q_COLIS', 'FTD_LIBELLE', 'DATE_LIMIT']
        if not all(column in df.columns for column in necessary_columns):
            messagebox.showwarning("Attention", f"Colonnes nécessaires manquantes dans {file_path}. Ce fichier sera ignoré.")
            return pd.DataFrame()  # Retourner un DataFrame vide
        else:
            # Filtrer les lignes contenant 'SOLDEUR' dans 'FTD_LIBELLE'
            df = df[df['FTD_LIBELLE'] == 'SOLDEUR']

            # Extraire la date du nom de fichier et l'ajouter comme colonne
            df['Semaine'] = extract_date_from_filename(file_path)

            # Formater la colonne DATE_LIMIT au format Jour/Mois/Année
            df['DATE_LIMIT'] = pd.to_datetime(df['DATE_LIMIT'], errors='coerce').dt.strftime('%d/%m/%Y')

            # Trier par 'INV_Q_COLIS' en ordre décroissant
            df.sort_values(by='INV_Q_COLIS', ascending=False, inplace=True)

            # Extraire le numéro de référence et le formater
            df['REF'] = df['REF & DESCRIPTION'].apply(lambda x: x.split(' - ')[0][1:-2])
            # Extraire le nom de la référence
            df['REF_NAME'] = df['REF & DESCRIPTION'].apply(lambda x: x.split(' - ')[1])

            # Sélectionner les colonnes pertinentes
            df = df[['REF', 'REF_NAME', 'INV_Q_COLIS', 'DATE_LIMIT', 'Semaine']]

            return df
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du traitement du fichier {file_path}: {e}")
        return pd.DataFrame()  # Retourner un DataFrame vide

# Fonction pour agréger les données par REF et DATE_LIMIT
def aggregate_data(df):
    # Supprimer les doublons en gardant celui avec le plus haut 'INV_Q_COLIS'
    df = df.sort_values('INV_Q_COLIS', ascending=False).drop_duplicates(subset=['REF', 'DATE_LIMIT'], keep='first')
    return df

# Fonction pour calculer les ventes des quatre semaines précédentes
def calculate_previous_month_sales(week, year, ref, sales_df):
    sales_volume = 0
    for i in range(1, 5):
        previous_week = week - i
        previous_year = year
        if previous_week <= 0:
            previous_week += 52
            previous_year -= 1
        week_year = f"{previous_week}.{previous_year}"
        weekly_sales = sales_df[(sales_df['Delivery Date'] == week_year) & (sales_df['REF'] == ref)]
        if not weekly_sales.empty:
            week_sales_volume = weekly_sales['Issued Volume'].sum()
            sales_volume += week_sales_volume
    return sales_volume

# Fonction principale
def main(year, week_range):
    try:
        start_week = int(week_range[0].replace('S', ''))  # Supprimer 'S' et convertir en int
        end_week = int(week_range[1].replace('S', ''))    # Supprimer 'S' et convertir en int
    except ValueError:
        messagebox.showerror("Erreur", "Format de plage de semaines invalide. Veuillez entrer les semaines au format 'SXX-SXX'.")
        return

    # Répertoire contenant les fichiers Excel
    directory = '/home/jingerdwarf/Documents/Crunch/5. Hackathon/Stock Inventaire OK/'

    # Initialiser un DataFrame vide pour stocker tous les éléments
    all_data_df = pd.DataFrame(columns=['REF', 'REF_NAME', 'INV_Q_COLIS', 'DATE_LIMIT', 'Semaine'])

    # Boucler à travers chaque semaine de la plage spécifiée
    for week in range(start_week, end_week + 1):
        file_name = f'Stock Inventaire du {year} S{week}.xlsx'
        file_path = os.path.join(directory, file_name)
        
        if os.path.exists(file_path):
            # Traiter le fichier et obtenir les données
            week_df = process_file(file_path)
            # Ajouter les données au DataFrame principal
            all_data_df = pd.concat([all_data_df, week_df])
        else:
            messagebox.showwarning("Attention", f"Fichier non trouvé : {file_path}")

    # Agréger les données par REF et DATE_LIMIT
    aggregated_df = aggregate_data(all_data_df)

    # Calculer le total INV_Q_COLIS pour chaque REF
    total_df = aggregated_df.groupby(['REF', 'REF_NAME'])['INV_Q_COLIS'].sum().reset_index()
    total_df.columns = ['REF', 'REF_NAME', 'Total_INV_Q_COLIS']

    # Convertir 'Total_INV_Q_COLIS' en numérique
    total_df['Total_INV_Q_COLIS'] = pd.to_numeric(total_df['Total_INV_Q_COLIS'], errors='coerce')

    # Calculer l'INV_Q_COLIS par mois pour chaque REF
    aggregated_df['Month'] = pd.to_datetime(aggregated_df['DATE_LIMIT'], format='%d/%m/%Y').dt.to_period('M')
    monthly_df = aggregated_df.groupby(['REF', 'REF_NAME', 'Month', 'Semaine'])['INV_Q_COLIS'].sum().reset_index()
    monthly_df.columns = ['REF', 'REF_NAME', 'Month', 'Semaine', 'Monthly_INV_Q_COLIS']

    # Fusionner total_df avec monthly_df
    result_df = pd.merge(total_df, monthly_df, on=['REF', 'REF_NAME'])
    
    # Lire les données de vente
    sales_file_path = '/home/jingerdwarf/Documents/Crunch/5. Hackathon/Datas pour les étudiants/Ventes 2023-2024.xlsx'
    if not os.path.exists(sales_file_path):
        messagebox.showerror("Erreur", f"Fichier de ventes non trouvé : {sales_file_path}")
        return

    # Lire la feuille 'Feuil1' du fichier des ventes
    sales_df = pd.read_excel(sales_file_path, sheet_name='Feuil1')
    
    if sales_df.empty:
        messagebox.showerror("Erreur", "Le fichier de ventes est vide. Veuillez vérifier le contenu du fichier.")
        return

    # Supprimer les espaces en début et fin de chaîne dans les noms de colonnes
    sales_df.columns = sales_df.columns.str.strip()

    # Renommer les colonnes pour la cohérence
    sales_df.rename(columns={
        'Unnamed: 10': 'REF',
        'Unnamed: 3': 'Plant',
        'Unnamed: 5': 'Y Cust Hier Level 6',
        'Unnamed: 8': 'Material',
        'Issued Volume': 'Issued Volume',
        'Delivery Date': 'Delivery Date'
    }, inplace=True)

    # Assurer que les colonnes 'REF' et 'Delivery Date' sont des chaînes de caractères
    sales_df['REF'] = sales_df['REF'].astype(str)
    sales_df['Delivery Date'] = sales_df['Delivery Date'].astype(str)

    # Convertir 'Issued Volume' en numérique, en forçant les erreurs
    sales_df['Issued Volume'] = pd.to_numeric(sales_df['Issued Volume'], errors='coerce')
    
    # Extraire le numéro de semaine de la colonne 'Semaine' pour le calcul correct
    result_df['Week_Number'] = result_df['Semaine'].apply(lambda x: int(x.split('.')[0]))

    # Calculer les ventes du mois précédent pour chaque article
    result_df['Previous Month Sales'] = result_df.apply(
        lambda row: calculate_previous_month_sales(row['Week_Number'], int(year), row['REF'], sales_df), axis=1
    )

    # Supprimer la colonne auxiliaire 'Week_Number'
    result_df.drop(columns=['Week_Number'], inplace=True)

    # Calculer le pourcentage de ventes par rapport à Monthly_INV_Q_COLIS
    result_df['Sales_Percentage'] = result_df.apply(
        lambda row: (row['Monthly_INV_Q_COLIS'] / row['Previous Month Sales'] * 100) if row['Previous Month Sales'] > 0 else 0,
        axis=1
    ).astype(int)

    # Exporter le DataFrame final dans un fichier CSV
    output_file_path = os.path.join(directory, f'aggregated_inventory_{year}_S{start_week}_S{end_week}.csv')
    result_df.to_csv(output_file_path, index=False)

    # Afficher le DataFrame final
    messagebox.showinfo("Résultat", f"Données agrégées enregistrées dans {output_file_path}")

    # Tracer le graphique en barres empilées pour les 10 premiers REF par Total_INV_Q_COLIS
    top_10_df = result_df.groupby(['REF', 'REF_NAME']).agg({'Total_INV_Q_COLIS': 'sum'}).nlargest(10, 'Total_INV_Q_COLIS').reset_index()
    top_10_data = result_df[result_df['REF'].isin(top_10_df['REF'])]

    fig, ax = plt.subplots(figsize=(15, 8))

    # Préparer les couleurs pour chaque mois unique
    months = top_10_data['Month'].unique()
    cmap = plt.cm.get_cmap('tab20', len(months))  # Utiliser un colormap avec plus de couleurs
    colors = [cmap(i) for i in range(len(months))]

    # Tracer les barres empilées
    bottom = pd.Series([0] * len(top_10_df['REF']), index=top_10_df['REF'])
    for i, month in enumerate(months):
        month_data = top_10_data[top_10_data['Month'] == month]
        ax.bar(month_data['REF'], month_data['Monthly_INV_Q_COLIS'], bottom=bottom[month_data['REF']], color=colors[i], label=month)
        bottom[month_data['REF']] += month_data['Monthly_INV_Q_COLIS'].values

    plt.title('Top 10 REF par Total_INV_Q_COLIS avec répartition mensuelle')
    plt.xlabel('REF')
    plt.ylabel('INV_Q_COLIS')
    plt.legend(title='Month')
    plt.tight_layout()

    # Enregistrer le graphique en tant que fichier image
    plot_output_path = os.path.join(directory, f'top_10_ref_inventory_{year}_S{start_week}_S{end_week}.png')
    plt.savefig(plot_output_path)

    # Afficher le graphique
    plt.show()

def run_app():
    # Créer la fenêtre principale
    root = Tk()
    root.title("Analyse d'Inventaire")

    # Ajouter les étiquettes et champs de saisie
    Label(root, text="Année").grid(row=0, column=0, padx=10, pady=10)
    year_entry = Entry(root)
    year_entry.grid(row=0, column=1, padx=10, pady=10)

    Label(root, text="Interval Semaines (format: SXX-SXX)").grid(row=1, column=0, padx=10, pady=10)
    week_entry = Entry(root)
    week_entry.grid(row=1, column=1, padx=10, pady=10)

    def on_submit():
        year = year_entry.get()
        week_range = week_entry.get().split('-')
        main(year, week_range)

    # Ajouter le bouton de soumission
    submit_button = Button(root, text="Soumettre", command=on_submit)
    submit_button.grid(row=2, column=0, columnspan=2, pady=20)

    # Lancer la boucle principale de l'interface graphique
    root.mainloop()

if __name__ == "__main__":
    run_app()
