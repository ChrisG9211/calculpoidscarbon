#! python3
# -*- coding: utf-8 -*-
if not list_of_results:
    from main_script import list_of_results, element_quantity_dict, category_quantity_dict, doc, path
from outil import round_3_decimals
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import datetime
import os 

# Define variables for plotly code
list_of_results = list_of_results
element_quantity_dict = element_quantity_dict
category_quantity_dict = category_quantity_dict
doc = doc
path = path
total = 0
mcrcom = None
prod = None
impact = None
sous_projet_totals = {}
material_totals = {}
material_par_sp = {}

for result in list_of_results:
    total_value = float(result["Impact sur le changement climatique (kgCO2e)"])

    # Convert total to a string if it's an integer
    if isinstance(total, int):
        total = str(total)

    # Remove commas
    total = total.replace(",", "")

    # Format the result with commas
    total = "{:,.3f}".format(float(total) + total_value)
    mcrcom = result["Macro-composant de niveau 1"]
    prod = result["Produit"]
    impact = round_3_decimals(
        result["Impact sur le changement climatique (kgCO2e)"]
    )

    if mcrcom not in sous_projet_totals:
        sous_projet_totals.update(
            {
                mcrcom: round_3_decimals(
                    result["Impact sur le changement climatique (kgCO2e)"]
                )
            }
        )
    else:
        sous_projet_totals[mcrcom] += round_3_decimals(
            result["Impact sur le changement climatique (kgCO2e)"]
        )

    if prod not in material_totals:
        material_totals.update({prod: impact})
    else:
        material_totals[prod] += impact

    if mcrcom not in material_par_sp:
        material_par_sp[mcrcom] = {prod: impact}
    else:
        material_par_sp[mcrcom][prod] = impact

fig = make_subplots(
    rows=2,
    cols=2,
    subplot_titles=(
        "Sous Projet Totals",
        "Materiaux Totals",
        "Materiel par Sous Projet",
        "Familles totals",
    ),
)

# Create bar charts for Sous Projet Totals
fig.add_trace(
    go.Bar(
        x=list(sous_projet_totals.keys()),
        y=list(sous_projet_totals.values()),
        name="Sous Projet Totals",
    ),
    row=1,
    col=1,
)

# Create bar charts for Material Totals
fig.add_trace(
    go.Bar(
        x=list(material_totals.keys()),
        y=list(material_totals.values()),
        name="Matériaux Totals",
    ),
    row=1,
    col=2,
)

# Create bar charts for Material by Sous Projet
for sp, material_data in material_par_sp.items():
    fig.add_trace(
        go.Bar(
            x=list(material_data.keys()),
            y=list(material_data.values()),
            name=f"{sp}",
        ),
        row=2,
        col=1,
    )

# Create bar charts for Material by Category
fig.add_trace(
    go.Bar(
        x=list(category_quantity_dict.keys()),
        y=list(category_quantity_dict.values()),
        name="Familles totals",
    ),
    row=2,
    col=2,
)

# Update x-axis labels for all subplots
fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=1)
fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=2)
fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=1)
fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=2)

moyen_total = total.replace(",", "")
Moyen_poids_carbone_par_element = "{:,.0f}".format(
    float(moyen_total) / len(element_quantity_dict)
)

# Update layout
fig.update_layout(
    barmode="group",
    title_text=f"Resultats poids carbone(kgCO2e): {total}\n    Projet: {doc.Title}\n    Nombre d'elements calcules: {'{:,.0f}'.format(len(element_quantity_dict))}    Moyen poids carbone par element(kgCO2e): {Moyen_poids_carbone_par_element}",
)
# Save the plot to an HTML file using the open function with 'utf-8' encoding
current_datetime = datetime.datetime.now().strftime("%y%m%d %Hh%M")

# Create file
html_file = "carbon_data {}.html".format(current_datetime)

# Write file at given directory recently created
with open(os.path.join(path, html_file), "w", encoding="utf-8") as file:
    file.write(fig.to_html())