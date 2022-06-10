"""
Created by Luca Rougeron
Modified by Maxime Vaillant
"""

import pandas as pd
import re
import sys


def get_top_10_article(data, article):
    aggregation = data[data.nom_article.eq(article)] \
        .groupby(by=["nom_article", "prenom_nom"]) \
        .agg({"quantite": "sum", "total": "sum"})

    return aggregation.sort_values(by=["quantite"], ascending=False).head(10)


def get_top_10_category(data, category):
    aggregation = data[data.famille_article.eq(category)] \
        .groupby(by=["famille_article", "prenom_nom"]) \
        .agg({"quantite": "sum", "total": "sum"})

    return aggregation.sort_values(by=["total"], ascending=False).head(10)


def export_top_category(data, categories, save_path):
    top_list = []
    for category in categories:
        with pd.ExcelWriter('{}/Top {}.xlsx'.format(save_path, category)) as writer:
            top_10_category = get_top_10_category(data, category)
            top_list.append(top_10_category)
            top_10_category.to_excel(writer, sheet_name=re.sub(r"[^a-zA-Z0-9]", "", category))

            data_category = data[data.famille_article.eq(category)]
            data_brut = data_category.groupby(by=["nom_article"]) \
                .agg({"quantite": "sum"}).reset_index().sort_values(by="quantite", ascending=False)
            list_article = data_brut.nom_article.to_list()

            for article in list_article:
                top_10_article = get_top_10_article(data, article)
                if top_10_article.iloc[0]['quantite'] > 10:
                    top_list.append(top_10_article)
                    top_10_article.to_excel(writer, sheet_name=re.sub(r"[^a-zA-Z0-9]", "", article))

    return top_list

def export_top(read_path, save_path):
    print('Loading data')

    data = pd.read_excel(read_path, engine="openpyxl")

    print('Data loaded')

    data['prenom_nom'] = data['Prénom acheteur'] + ' ' + data['Nom acheteur']
    data = data[['prenom_nom', 'Article', 'Famille d\'article', 'Quantité', 'Total TTC']]
    data.rename(columns={
        'Article': 'nom_article',
        'Famille d\'article': 'famille_article',
        'Quantité': 'quantite',
        'Total TTC': 'total'
    }, inplace=True)

    TOP_CATEGORY = [
        "Bières pression",
        "Bières bouteilles",
        "Softs",
        "Repas",
        "Vrac",
        "Pampryl",
        "Petit Dej"
    ]

    top_list = export_top_category(data, TOP_CATEGORY, save_path)

    # Extraction du top des tops
    COEF_TOP_TOP = [10, 9, 8, 7, 6, 5, 4, 3, 2, 1]

    clean_top_list = []
    for top in top_list:
        clean_top = top.reset_index()
        clean_top['coef'] = COEF_TOP_TOP[:len(clean_top['prenom_nom'])]
        clean_top_list.append(clean_top)

    top_tops = pd.concat(clean_top_list)
    top_tops.groupby(['prenom_nom']).agg({'coef': 'sum'}) \
        .sort_values(by=["coef"], ascending=False).head(10) \
        .to_excel('{}/Top Top.xlsx'.format(save_path), sheet_name='Top top')

    # Extraction du top dépense
    data.groupby(by=["prenom_nom"]) \
        .agg({"total": "sum"}).sort_values(by="total", ascending=False).head(10) \
        .to_excel('{}/Top Dépense.xlsx'.format(save_path), sheet_name='Dépense')

    print('Export done')


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        export_top(sys.argv[1], sys.argv[2])
    else:
        print('Les arguments ne sont pas bons')
