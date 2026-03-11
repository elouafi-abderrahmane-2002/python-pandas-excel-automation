# 📋 Python + pandas + Excel — Automatisation de l'Analyse de Données

Recevoir chaque semaine un fichier Excel de 10 000 lignes, le nettoyer manuellement,
calculer des agrégations et envoyer un rapport formaté — c'est une tâche répétitive
qu'on peut automatiser en quelques dizaines de lignes de Python.

Ce projet est une boîte à outils pour faire exactement ça : lire des fichiers Excel
complexes, les transformer avec pandas, et générer des rapports Excel formatés
automatiquement avec openpyxl.

---

## Ce que fait ce projet

```
  Fichier(s) Excel en entrée
  (bruts, mal formatés, multi-feuilles)
          │
          │  Lecture intelligente
          ▼
  ┌────────────────────────────┐
  │       pandas               │
  │                            │
  │  pd.read_excel()           │
  │  - Sélection de feuilles   │
  │  - Skiprows si besoin      │
  │  - Types de colonnes       │
  └────────────┬───────────────┘
               │
               │  Nettoyage & transformation
               ▼
  ┌────────────────────────────┐
  │  Pipeline de nettoyage     │
  │                            │
  │  - Supprimer doublons      │
  │  - Gérer NaN (fill/drop)   │
  │  - Normaliser les formats  │
  │  - Convertir les types     │
  │  - Valider les plages      │
  └────────────┬───────────────┘
               │
               │  Analyse & agrégations
               ▼
  ┌────────────────────────────┐
  │  pandas groupby / pivot    │
  │                            │
  │  - KPIs par catégorie      │
  │  - Tendances temporelles   │
  │  - Tableaux croisés        │
  │  - Détection d'anomalies   │
  └────────────┬───────────────┘
               │
               │  Génération du rapport
               ▼
  ┌────────────────────────────┐
  │  openpyxl                  │
  │                            │
  │  - Mise en forme (couleurs │
  │    polices, bordures)      │
  │  - Graphiques intégrés     │
  │  - Formules Excel          │
  │  - Protection de feuilles  │
  └────────────────────────────┘
               │
               ▼
  rapport_final_YYYYMMDD.xlsx
  (prêt à envoyer par email)
```

---

## Lecture & nettoyage

```python
import pandas as pd
import numpy as np

# Lire un Excel avec plusieurs feuilles
df_sales   = pd.read_excel('data_raw.xlsx', sheet_name='Sales',
                            skiprows=2, dtype={'ID': str})
df_clients = pd.read_excel('data_raw.xlsx', sheet_name='Clients')

# Nettoyage
def clean_dataframe(df):
    df = df.drop_duplicates()
    df = df.dropna(subset=['ID', 'Amount'])    # colonnes critiques
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df['Date']   = pd.to_datetime(df['Date'],  errors='coerce')
    df['Region'] = df['Region'].str.strip().str.title()
    return df

df_sales   = clean_dataframe(df_sales)
df_clients = clean_dataframe(df_clients)

# Jointure
df = df_sales.merge(df_clients, on='ClientID', how='left')
```

---

## Agrégations & KPIs

```python
# Ventes par région et par mois
monthly = (df.groupby([df['Date'].dt.to_period('M'), 'Region'])
             .agg(
                 Revenue=('Amount', 'sum'),
                 Orders=('ID', 'count'),
                 AvgOrder=('Amount', 'mean')
             )
             .reset_index())

# Détection d'anomalies (Z-score)
from scipy import stats
df['z_score'] = np.abs(stats.zscore(df['Amount'].fillna(0)))
anomalies = df[df['z_score'] > 3]
print(f"{len(anomalies)} transactions anormales détectées")

# Tableau croisé dynamique
pivot = df.pivot_table(
    values='Amount',
    index='Region',
    columns=df['Date'].dt.month_name(),
    aggfunc='sum',
    fill_value=0,
    margins=True
)
```

---

## Génération du rapport Excel formaté

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

wb = Workbook()
ws = wb.active
ws.title = "Rapport Mensuel"

# En-tête stylisé
header_fill = PatternFill("solid", fgColor="2F5597")  # bleu foncé
header_font = Font(bold=True, color="FFFFFF", size=11)

for col_num, col_name in enumerate(monthly.columns, 1):
    cell = ws.cell(row=1, column=col_num, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center')

# Données avec alternance de couleurs
alt_fill = PatternFill("solid", fgColor="DCE6F1")
for row_num, row in enumerate(dataframe_to_rows(monthly, index=False, header=False), 2):
    for col_num, value in enumerate(row, 1):
        cell = ws.cell(row=row_num, column=col_num, value=value)
        if row_num % 2 == 0:
            cell.fill = alt_fill

# Ajuster la largeur des colonnes automatiquement
for col in ws.columns:
    max_length = max(len(str(cell.value or '')) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_length + 4

# Ajouter un graphique à barres
chart = BarChart()
chart.title = "Revenue par Région"
data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)
ws.add_chart(chart, "H2")

# Sauvegarder
from datetime import date
filename = f"rapport_{date.today().strftime('%Y%m%d')}.xlsx"
wb.save(filename)
print(f"Rapport généré : {filename}")
```

---

## Ce que j'ai vraiment appris

La distinction entre `pd.read_excel()` et `openpyxl.load_workbook()` est importante.
pandas est fait pour *analyser* des données — il lit vite mais ignore tout le formatage.
openpyxl permet de *créer et modifier* des fichiers Excel avec un contrôle total
sur le style, les formules, et les graphiques — mais c'est plus verbeux.

Pour les rapports professionnels, la combinaison est idéale : pandas pour les calculs,
openpyxl pour la présentation finale. C'est exactement ce que font les vrais outils
de reporting en entreprise, juste sans interface graphique.

---

*Projet réalisé dans le cadre de ma formation ingénieur — ENSET Mohammedia*
*Par **Abderrahmane Elouafi** · [LinkedIn](https://www.linkedin.com/in/abderrahmane-elouafi-43226736b/) · [Portfolio](https://my-first-porfolio-six.vercel.app/)*
