# Import the necessary libraries
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import io






        
###############################
#   Function Definitions
###############################

font_styles = None  # sera rempli dans main()

def get_font_styles():
    st.sidebar.subheader("🖋️ Font Style Settings")

    family = st.sidebar.selectbox("Font Family", ["DejaVu Sans", "sans-serif", "serif"], index=0)
    fontweight = st.sidebar.selectbox("Font Weight", ["normal", "bold"], index=0)
    fontstyle = st.sidebar.selectbox("Font Style", ["normal", "italic"], index=0)

    title_size = st.sidebar.slider("Title Font Size", 10, 30, 16)
    label_size = st.sidebar.slider("Label Font Size", 8, 30, 12)
    text_size = st.sidebar.slider("Text Font Size", 6, 30, 10)
    legend_size = st.sidebar.slider("Legend Font Size", 6, 30, 10)

    return {
        "family": family,
        "weight": fontweight,
        "style": fontstyle,
        "title_size": title_size,
        "label_size": label_size,
        "text_size": text_size,
        "legend_size": legend_size
    }


def analyze_excel_and_generate_tables(file_path, sheet_name=0):
    """
    Analyzes the specified Excel file and sheet, identifying scenarios, contributions,
    and generating initial and combined tables for impact assessment.
    """
    try:
        # Read the Excel file
        data = pd.read_excel(file_path, header=None, sheet_name=sheet_name)
        print(f"Excel file loaded successfully from sheet: {sheet_name}!")

        # Step 1: Detect scenarios
        scenario_row = data.iloc[0, 1:]
        scenario_names = scenario_row.unique()
        scenario_names_cleaned = [
            name[name.find("(") + 1 : name.find(")")].strip().lower()
            for name in scenario_names
        ]
        print(f"Detected scenarios ({len(scenario_names_cleaned)}): {scenario_names_cleaned}")

        # Step 2: Detect contributions
        contribution_row = data.iloc[1, 1:]
        print(f"Detected contributions: {list(contribution_row.unique())}")

        # Step 3: Structure the data
        categories = data.iloc[2:, 0]
        values = data.iloc[2:, 1:]
        values.columns = pd.MultiIndex.from_arrays([scenario_row, contribution_row])
        values.index = categories

        # Initial table
        initial_table = values.copy()

        # Impact totals (absolute values)
        total_impact_table = values.abs().T.groupby(level=0).sum().T

        # Replace columns with simplified scenario names
        total_impact_table.columns = [
            name[name.find("(") + 1 : name.find(")")].strip()
            for name in total_impact_table.columns
        ]

        # Relative percentages
        relative_percentage_table = total_impact_table.copy()
        for category in relative_percentage_table.index:
            max_value = relative_percentage_table.loc[category].max()
            relative_percentage_table.loc[category] = (
                relative_percentage_table.loc[category] / max_value
            ) * 100
        relative_percentage_table = relative_percentage_table.add_suffix(" (%)")

        # Combine tables
        combined_table = pd.concat([total_impact_table, relative_percentage_table], axis=1)

        return initial_table, combined_table, total_impact_table, scenario_names_cleaned
    except Exception as e:
        print(f"Error analyzing the Excel file: {e}")
        return None, None, None, None

def extract_contributions_from_initial_table(initial_table):
    """
    Extrait les contributions uniques depuis les colonnes MultiIndex de initial_table.
    """
    try:
        if not isinstance(initial_table.columns, pd.MultiIndex):
            raise ValueError("initial_table n'a pas un MultiIndex en colonnes.")

        contributions = initial_table.columns.get_level_values(1)
        unique_contributions = sorted(set(contributions.astype(str).str.strip()))

        print(f"Contributions uniques détectées ({len(unique_contributions)}) : {unique_contributions}")
        return unique_contributions

    except Exception as e:
        print(f"Erreur lors de l'extraction depuis initial_table : {e}")
        return []
    
def generate_percentage_table(initial_table, total_impact_table):
    """
    Generates a DataFrame identical to initial_table, but converts its values to percentages
    relative to the maximum impact found in total_impact_table for each category.
    """
    try:
        percentage_table = initial_table.copy()
        
        for category in initial_table.index:
            # Get the maximum impact for this category
            max_total_impact = total_impact_table.loc[category].max()
            # Calculate percentages
            percentage_table.loc[category] = (
                initial_table.loc[category] / max_total_impact
            ) * 100

        return percentage_table
    except Exception as e:
        print(f"Error generating percentage table: {e}")
        return None


def generate_tables_by_scenario(initial_table, scenario_names_cleaned):
    """
    Creates a dictionary of tables, one per scenario, containing:
    - Impact categories
    - Contributions
    - Total impact per category
    - Percentage of each contribution relative to the total impact
    """
    try:
        scenario_tables = {}
        for scenario_clean in scenario_names_cleaned:
            # Find the matching scenario in the MultiIndex columns
            matching_scenario = [
                col for col in initial_table.columns.levels[0] if scenario_clean in col.lower()
            ]
            if not matching_scenario:
                raise KeyError(f"No column match found for simplified scenario '{scenario_clean}'.")
            matching_scenario = matching_scenario[0]

            # Filter data for this scenario
            scenario_data = initial_table.xs(key=matching_scenario, axis=1, level=0)

            # Create a formatted table
            scenario_table = pd.DataFrame()
            scenario_table["Impact Category"] = initial_table.index
            for contribution in scenario_data.columns:
                scenario_table[contribution] = scenario_data[contribution].values

            # Calculate total impact (absolute values)
            scenario_table["Total Impact"] = scenario_table.iloc[:, 1:].abs().sum(axis=1)

            # Calculate contribution percentages
            for contribution in scenario_data.columns:
                scenario_table[f"% {contribution}"] = (
                    scenario_table[contribution] / scenario_table["Total Impact"] * 100
                )

            scenario_tables[scenario_clean] = scenario_table

            print(f"\nGenerated table for scenario: {scenario_clean}")
            print(scenario_table.head())

        return scenario_tables
    except Exception as e:
        print(f"Error generating scenario tables: {e}")
        return None


def plot_comparison_bar_chart(total_impact_table):
    """
    Plots 3 figures:
    1) Main comparison bar chart
    2) Separate legend
    3) Bar chart with total values labeled
    Returns a list of matplotlib figures.
    """
    try:
        global scenario_colors, font_styles  # <== pour accéder aux styles depuis main()

        # Extraction des styles personnalisés
        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        figs = []

        # Normalize data
        normalized_data = total_impact_table.copy()
        for category in normalized_data.index:
            max_value = normalized_data.loc[category].max()
            normalized_data.loc[category] = (normalized_data.loc[category] / max_value) * 100

        categories = normalized_data.index
        scenarios = normalized_data.columns
        num_scenarios = len(scenarios)
        bar_width = 0.9 / num_scenarios
        x_positions = np.arange(len(categories))

        # Handle scenario colors
        scenario_color_map = None
        generic_colors = plt.get_cmap("tab10").colors

        if scenario_colors is not None:
            if isinstance(scenario_colors, dict):
                scenario_color_map = scenario_colors
            elif hasattr(scenario_colors, "iterrows"):
                scenario_color_map = {
                    row["Scenario"].strip().lower(): row["Hex Code"]
                    for _, row in scenario_colors.iterrows()
                }

        # Figure 1: Main bar chart (no totals)
        fig1, ax1 = plt.subplots(figsize=(12, 7), dpi=300)
        bars = []
        labels = []

        for i, scenario in enumerate(scenarios):
            key = scenario.strip().lower()
            color = (
                scenario_color_map.get(key, "#000000")
                if scenario_color_map else generic_colors[i % len(generic_colors)]
            )
            bar = ax1.bar(
                x_positions + i * bar_width,
                normalized_data[scenario],
                bar_width,
                color=color,
            )
            bars.append(bar[0])
            labels.append(scenario)

        ax1.set_title("Comparison of Scenarios", fontsize=title_size, fontweight=fontweight,
              fontstyle=fontstyle, family=fontfamily, pad=30)

        ax1.set_ylabel("(%)", fontsize=label_size, fontweight=fontweight,
                       fontstyle=fontstyle, family=fontfamily)
        ax1.set_ylim(0, 100)
        ax1.set_xlim(-0.5, len(categories))
        ax1.set_xticks(x_positions + (len(scenarios) - 1) * bar_width / 2)
        ax1.set_xticklabels(categories, rotation=45, ha="right", fontsize=label_size,
                            fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
        ax1.grid(axis="y", linestyle="--", alpha=0.7)
        ax1.tick_params(axis='y', labelsize=label_size)

        plt.tight_layout()
        figs.append(fig1)
        plt.close(fig1)

        # Figure 2: Legend
        fig2, ax2 = plt.subplots(figsize=(5, 3), dpi=300)
        ax2.axis("off")
        ax2.legend(bars, labels, title="Scenarios", fontsize=legend_size, title_fontsize=legend_size, loc="center")
        plt.tight_layout()
        figs.append(fig2)
        plt.close(fig2)

        # Figure 3: With totals
        fig3, ax3 = plt.subplots(figsize=(12, 7), dpi=300)
        for i, scenario in enumerate(scenarios):
            key = scenario.strip().lower()
            color = (
                scenario_color_map.get(key, "#000000")
                if scenario_color_map else generic_colors[i % len(generic_colors)]
            )
            ax3.bar(
                x_positions + i * bar_width,
                normalized_data[scenario],
                bar_width,
                color=color,
            )

        for i, scenario in enumerate(scenarios):
            for j, value in enumerate(normalized_data[scenario]):
                total_value = total_impact_table.loc[categories[j], scenario]
                formatted_value = f"{total_value:.2f}" if 0.01 <= abs(total_value) <= 1000 else f"{total_value:.2E}"
                ax3.text(
                    x_positions[j] + i * bar_width,
                    value + 2,
                    formatted_value,
                    ha="center",
                    va="bottom",
                    fontsize=text_size,
                    color="black",
                    fontweight=fontweight,
                    fontstyle=fontstyle,
                    family=fontfamily,
                    rotation=90,
                )

        ax3.set_title("Comparison of Scenarios (with Totals)", fontsize=title_size, fontweight=fontweight,
                      fontstyle=fontstyle, family=fontfamily, pad=80)
        ax3.set_ylabel("(%)", fontsize=label_size, fontweight=fontweight,
                       fontstyle=fontstyle, family=fontfamily)
        ax3.set_ylim(0, 100)
        ax3.set_xlim(-0.5, len(categories))
        ax3.set_xticks(x_positions + (len(scenarios) - 1) * bar_width / 2)
        ax3.set_xticklabels(categories, rotation=45, ha="right", fontsize=label_size,
                            fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
        ax3.grid(axis="y", linestyle="--", alpha=0.7)
        ax3.tick_params(axis='y', labelsize=label_size)

        plt.tight_layout()
        figs.append(fig3)
        plt.close(fig3)

        return figs

    except Exception as e:
        st.error(f"Error generating comparison bar chart: {str(e)}")
        return []




def plot_relative_contribution_by_scenario(scenario_tables, contributions_order):

    """
    Génère une liste de (main_figure, legend_figure) pour chaque scénario,
    affichant un graphique en barres empilées (verticales) des contributions relatives.
    """
    try:
        global contributions_colors, font_styles

        # 🔠 Extraction du style de police
        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        figures = []

        # Identifier toutes les contributions présentes
        all_contributions = set()
        for table in scenario_tables.values():
            contribution_columns = [f"% {c}" for c in contributions_order if f"% {c}" in table.columns]

            all_contributions.update(contribution_columns)

        # Générer la palette de couleurs
        contribution_colors = {}
        generic_colors = plt.get_cmap("tab10").colors

        if contributions_colors is not None:
            if isinstance(contributions_colors, dict):
                for i, contrib in enumerate(sorted(all_contributions)):
                    key = contrib.replace("%", "").strip().lower()
                    contribution_colors[contrib] = contributions_colors.get(key, generic_colors[i % len(generic_colors)])
            elif hasattr(contributions_colors, "iterrows"):
                for i, contrib in enumerate(sorted(all_contributions)):
                    contrib_clean = contrib.replace("%", "").strip().lower()
                    match = contributions_colors[
                        contributions_colors["Contributions"].str.lower() == contrib_clean
                    ]
                    if not match.empty:
                        contribution_colors[contrib] = match["Hex Code"].iloc[0]
                    else:
                        contribution_colors[contrib] = generic_colors[i % len(generic_colors)]
        else:
            for i, contrib in enumerate(sorted(all_contributions)):
                contribution_colors[contrib] = generic_colors[i % len(generic_colors)]

        # Génération des graphiques
        for scenario_name, table in scenario_tables.items():
            main_fig, main_ax = plt.subplots(figsize=(16, 8), dpi=300)
            legend_fig, legend_ax = plt.subplots(figsize=(5, 3), dpi=300)
            
            categories = table["Impact Category"]
            contribution_columns = [
                col for col in table.columns
                if col.startswith("%") and col != "%Total Impact"
            ]
            percentage_data = table[contribution_columns].astype(float)
            total_impact = table["Total Impact"].astype(float)
            bar_width = 0.6
            bars = []
            labels = []
            
            bottom_pos = np.zeros(len(categories))
            bottom_neg = np.zeros(len(categories))

            for contribution in contribution_columns:
                values = percentage_data[contribution].values
                pos_values = np.where(values > 0, values, 0)
                neg_values = np.where(values < 0, values, 0)

                if np.any(pos_values):
                    bar = main_ax.bar(
                        categories, pos_values, bar_width,
                        bottom=bottom_pos,
                        color=contribution_colors[contribution]
                    )
                    bars.append(bar[0])
                    labels.append(contribution.replace("%", ""))
                    bottom_pos += pos_values

                if np.any(neg_values):
                    main_ax.bar(
                        categories, neg_values, bar_width,
                        bottom=bottom_neg,
                        color=contribution_colors[contribution]
                    )
                    bottom_neg += neg_values

            main_ax.axhline(0, color="black", linestyle="--", alpha=0.7)

            # Ajouter les valeurs totales au-dessus
            for i, total in enumerate(total_impact):
                formatted_value = f"{total:.2f}" if abs(total) <= 1000 else f"{total:.2E}"
                main_ax.text(
                    i, max(bottom_pos[i], 0) + 1,
                    formatted_value,
                    ha="center",
                    va="bottom",
                    fontsize=text_size,
                    color="black",
                    fontweight=fontweight,
                    fontstyle=fontstyle,
                    family=fontfamily
                )

            main_ax.set_title(
                f"Relative Contributions - {scenario_name.capitalize()}",
                fontsize=title_size,
                pad=40,
                fontweight=fontweight,
                fontstyle=fontstyle,
                family=fontfamily
            )
            main_ax.set_ylabel("(%)", fontsize=label_size, fontweight=fontweight,
                               fontstyle=fontstyle, family=fontfamily)
            main_ax.set_ylim(min(bottom_neg) - 5, max(bottom_pos))
            main_ax.set_xticks(range(len(categories)))
            main_ax.set_xticklabels(
                categories,
                rotation=45,
                ha="right",
                fontsize=label_size,
                fontweight=fontweight,
                fontstyle=fontstyle,
                family=fontfamily
            )
            main_ax.grid(axis="y", linestyle="--", alpha=0.7)
            main_ax.tick_params(axis='y', labelsize=label_size)
            plt.tight_layout()

            legend_ax.axis("off")
            legend_ax.legend(bars, labels, title="Contributions", fontsize=legend_size, title_fontsize=legend_size, loc="center")
            plt.tight_layout()

            figures.append((main_fig, legend_fig))
            

        return figures

    except Exception as e:
        st.error(f"Error generating relative contributions: {str(e)}")
        return []





def plot_relative_contribution_by_scenario_horizontal(scenario_tables, contributions_order):

    """
    Génère des (main_figure, legend_figure) pour chaque scénario
    affichant un bar chart horizontal empilé des contributions relatives.
    """
    try:
        global contributions_colors, font_styles

        # 🔠 Styles de police
        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        figures = []

        # Identifier toutes les contributions
        all_contributions = set()
        for table in scenario_tables.values():
            
            contribution_columns = [f"% {c}" for c in contributions_order if f"% {c}" in table.columns]
                
            all_contributions.update(contribution_columns)

        # Générer la palette de couleurs
        contribution_color_map = {}
        generic_colors = plt.get_cmap("tab10").colors

        if contributions_colors is not None:
            if isinstance(contributions_colors, dict):
                for i, contrib in enumerate(sorted(all_contributions)):
                    key = contrib.replace("%", "").strip().lower()
                    contribution_color_map[contrib] = contributions_colors.get(key, generic_colors[i % len(generic_colors)])
            elif hasattr(contributions_colors, "iterrows"):
                for i, contrib in enumerate(sorted(all_contributions)):
                    contrib_clean = contrib.replace("%", "").strip().lower()
                    match = contributions_colors[
                        contributions_colors["Contributions"].str.lower() == contrib_clean
                    ]
                    if not match.empty:
                        contribution_color_map[contrib] = match["Hex Code"].iloc[0]
                    else:
                        contribution_color_map[contrib] = generic_colors[i % len(generic_colors)]
        else:
            for i, contrib in enumerate(sorted(all_contributions)):
                contribution_color_map[contrib] = generic_colors[i % len(generic_colors)]

        # Génération des graphiques
        for scenario_name, table in scenario_tables.items():
            main_fig, main_ax = plt.subplots(figsize=(10, 12), dpi=300)
            legend_fig, legend_ax = plt.subplots(figsize=(5, 3), dpi=300)

            categories = table["Impact Category"]
            contribution_columns = [
                col for col in table.columns
                if col.startswith("%") and col != "%Total Impact"
            ]
            percentage_data = table[contribution_columns].astype(float)
            total_impact = table["Total Impact"].astype(float)
            bar_height = 0.6
            bars = []
            labels = []
            
            left_pos = np.zeros(len(categories))
            left_neg = np.zeros(len(categories))

            for contribution in contribution_columns:
                values = percentage_data[contribution].values
                pos_values = np.where(values > 0, values, 0)
                neg_values = np.where(values < 0, values, 0)

                if np.any(pos_values):
                    bar = main_ax.barh(
                        categories,
                        pos_values,
                        bar_height,
                        left=left_pos,
                        color=contribution_color_map[contribution]
                    )
                    bars.append(bar[0])
                    labels.append(contribution.replace("%", ""))
                    left_pos += pos_values

                if np.any(neg_values):
                    main_ax.barh(
                        categories,
                        neg_values,
                        bar_height,
                        left=left_neg,
                        color=contribution_color_map[contribution]
                    )
                    left_neg += neg_values

            main_ax.axvline(0, color="black", linestyle="--", alpha=0.7)

            # Affichage des valeurs totales à droite
            for i, total in enumerate(total_impact):
                formatted_value = f"{total:.2f}" if abs(total) <= 1000 else f"{total:.2E}"
                main_ax.text(
                    max(left_pos[i], 0) + 5,
                    i,
                    formatted_value,
                    ha="left",
                    va="center",
                    fontsize=text_size,
                    color="black",
                    fontweight=fontweight,
                    fontstyle=fontstyle,
                    family=fontfamily
                )

            main_ax.set_title(
                f"Relative Contributions - {scenario_name.capitalize()}",
                fontsize=title_size,
                pad=30,
                fontweight=fontweight,
                fontstyle=fontstyle,
                family=fontfamily
            )
            main_ax.set_xlabel("(%)", fontsize=label_size,
                               fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            main_ax.set_xlim(min(left_neg) - 5, max(left_pos))
            main_ax.set_yticks(range(len(categories)))
            main_ax.set_yticklabels(
                categories,
                fontsize=label_size,
                fontweight=fontweight,
                fontstyle=fontstyle,
                family=fontfamily
            )
            main_ax.grid(axis="x", linestyle="--", alpha=0.7)
            main_ax.tick_params(axis='x', labelsize=label_size)
            plt.tight_layout()

            legend_ax.axis("off")
            legend_ax.legend(bars, labels, title="Contributions",
                             fontsize=legend_size,title_fontsize=legend_size,  loc="center")
            plt.tight_layout()

            figures.append((main_fig, legend_fig))
           

        return figures

    except Exception as e:
        st.error(f"Error generating horizontal contributions: {str(e)}")
        return []




def plot_stacked_bar_by_category(initial_table, total_impact_table, contributions_order):

    """
    Génère (main_figure, legend_figure, category_name) pour chaque catégorie d’impact.
    Utilise les valeurs réelles (initial_table). Adapté pour affichage dans Streamlit via st.pyplot().
    """
    try:
        global contributions_colors, font_styles

        # 🔠 Styles de police
        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        figures = []

        if initial_table.empty or total_impact_table.empty:
            st.warning("⚠ Données vides pour les graphiques par catégorie.")
            return []

        # 🔹 Obtenir les scénarios et contributions
        scenarios = initial_table.columns.get_level_values(0).unique()
        contributions = initial_table.columns.get_level_values(1).unique()
        categories = initial_table.index

        # 🔹 Nettoyage des noms de scénarios
        scenario_labels = {scenario: scenario.split("(")[-1].replace(")", "").strip().lower() for scenario in scenarios}
        total_impact_labels = {sc.lower(): sc for sc in total_impact_table.columns}

        # 🔹 Générer le mapping de couleurs des contributions
        contrib_color_map = {}
        generic_colors = plt.get_cmap("tab10").colors
        if contributions_colors is not None:
            if isinstance(contributions_colors, dict):
                for i, contrib in enumerate(contributions):
                    key = contrib.replace("%", "").strip().lower()
                    contrib_color_map[contrib] = contributions_colors.get(key, generic_colors[i % len(generic_colors)])
            elif hasattr(contributions_colors, "iterrows"):
                for i, contrib in enumerate(contributions):
                    contrib_clean = contrib.replace("%", "").strip().lower()
                    match = contributions_colors[contributions_colors["Contributions"].str.lower() == contrib_clean]
                    if not match.empty:
                        contrib_color_map[contrib] = match["Hex Code"].iloc[0]
                    else:
                        contrib_color_map[contrib] = generic_colors[i % len(generic_colors)]
        else:
            for i, contrib in enumerate(contributions):
                contrib_color_map[contrib] = generic_colors[i % len(generic_colors)]

        # 🔹 Générer les graphiques pour chaque catégorie
        nombre_scenario=len(scenarios)
        for category in categories:
            category_data = initial_table.loc[category]
            x_positions = np.arange(len(scenarios))*0.7
            bar_width = 0.5 # réduit l'espace entre les barres

            main_fig, main_ax = plt.subplots(figsize=(nombre_scenario*0.9, 4), dpi=300)
            legend_fig, legend_ax = plt.subplots(figsize=(5, 3), dpi=300)

            bottom_pos = np.zeros(len(scenarios))
            bottom_neg = np.zeros(len(scenarios))
            bars = []
            labels = []

            for contribution in contributions_order:
                if contribution not in contributions:
                    continue  # ignorer si non trouvé
                values = []
                for s in scenarios:
                    try:
                        val = category_data[s, contribution]
                    except KeyError:
                        val = 0
                    values.append(val)
                values = np.array(values, dtype=float)

                pos_values = np.where(values > 0, values, 0)
                neg_values = np.where(values < 0, values, 0)

                if np.any(pos_values):
                    bar = main_ax.bar(
                        x_positions,
                        pos_values,
                        width=bar_width,
                        bottom=bottom_pos,
                        color=contrib_color_map.get(contribution, "#000000")
                    )
                    bars.append(bar[0])
                    labels.append(contribution.replace("%", ""))
                    bottom_pos += pos_values

                if np.any(neg_values):
                    main_ax.bar(
                        x_positions,
                        neg_values,
                        width=bar_width,
                        bottom=bottom_neg,
                        color=contrib_color_map.get(contribution, "#000000")
                    )
                    bottom_neg += neg_values

            # ✅ Calcul de l'échelle optimisée après empilement
            all_stacked_values = list(bottom_pos) + list(bottom_neg)
            min_val = min(all_stacked_values + [0])
            max_val = max(all_stacked_values + [0])
            y_margin = 0.05 * (max_val - min_val) if max_val != min_val else 1
            ylim_min = min_val - y_margin
            ylim_max = max_val + y_margin
            main_ax.set_ylim(ylim_min, ylim_max)

            # 🔹 Ajouter les totaux
            for i, scenario in enumerate(scenarios):
                scen_clean = scenario_labels[scenario]
                if scen_clean in total_impact_labels:
                    try:
                        total_value = total_impact_table.loc[category, total_impact_labels[scen_clean]]
                        formatted = f"{total_value:.2f}" if 0.01 <= abs(total_value) <= 1000 else f"{total_value:.2E}"
                        main_ax.text(
                            x_positions[i],
                            max(bottom_pos[i], 0) + y_margin * 0.3,
                            formatted,
                            ha='center',
                            va='bottom',
                            fontsize=text_size,
                            color='black',
                            fontweight=fontweight,
                            fontstyle=fontstyle,
                            family=fontfamily
                        )
                    except KeyError:
                        continue

            # 🔹 Mise en forme
    
            main_ax.set_ylabel(f"{category}", fontsize=label_size,
                               fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            main_ax.set_xticks(x_positions)
            main_ax.set_xticklabels(
                [scenario_labels[s] for s in scenarios],
                rotation=45,
                ha="right",
                fontsize=label_size,
                fontweight=fontweight,
                fontstyle=fontstyle,
                family=fontfamily
            )

            main_ax.set_xlim(x_positions[0] - (bar_width / 2), x_positions[-1] + (bar_width / 2))
            main_ax.grid(axis="y", linestyle="--", alpha=0.7)
            main_ax.tick_params(axis='y', labelsize=label_size)
            plt.tight_layout()

            legend_ax.axis("off")
            legend_ax.legend(bars, labels, title="Contributions",
                             fontsize=legend_size, title_fontsize=legend_size, loc="center")
            plt.tight_layout()

            figures.append((main_fig, legend_fig, category))

        return figures

    except Exception as e:
        st.error(f"❌ Error generating graph per category : {str(e)}")
        return []






def plot_combined_graph_with_scenario_hatches(percentage_table, total_impact_table, contributions_order):


    """
    Génère des graphiques combinés par scénario avec hachures et légendes.
    Retourne une liste de tuples : (description, figure).
    """
    try:
        global contributions_colors, scenario_colors, font_styles
        figures = []

        # Récupération des styles de police
        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        if percentage_table.empty or total_impact_table.empty:
            st.warning("Missing data for the combined scenario chart.")
            return []

        categories = percentage_table.index
        scenarios = percentage_table.columns.levels[0]
        contributions = percentage_table.columns.levels[1]
        num_scenarios = len(scenarios)
        bar_width = 0.9 / num_scenarios
        x_positions = np.arange(len(categories))*1.5

        # Définir les hachures par scénario
        scenario_hatch_dict = {}
        for scenario in scenarios:
              key = scenario.split("(")[-1].replace(")", "").strip().lower()
              scenario_hatch_dict[scenario] = scenario_hatches.get(key, None)


        # Gérer les couleurs des contributions
        cmap = plt.get_cmap("tab10")
        contrib_color_map = {}
        if contributions_colors is not None:
            if isinstance(contributions_colors, dict):
                for i, contrib in enumerate(contributions):
                    key = contrib.replace("%", "").strip().lower()
                    contrib_color_map[contrib] = contributions_colors.get(key, cmap(i % 10))
            elif hasattr(contributions_colors, "iterrows"):
                for i, contrib in enumerate(contributions):
                    key = contrib.replace("%", "").strip().lower()
                    match = contributions_colors[
                        contributions_colors["Contributions"].str.lower() == key
                    ]
                    contrib_color_map[contrib] = match["Hex Code"].iloc[0] if not match.empty else cmap(i % 10)
        else:
            for i, contrib in enumerate(contributions):
                contrib_color_map[contrib] = cmap(i % 10)

        # Gérer les couleurs des scénarios
        scenario_color_map = {}
        if scenario_colors is not None:
            if isinstance(scenario_colors, dict):
                for scen in scenarios:
                    key = scen.strip().lower()
                    scenario_color_map[scen] = scenario_colors.get(key, "#FFFFFF")
            elif hasattr(scenario_colors, "iterrows"):
                for scen in scenarios:
                    key = scen.strip().lower()
                    match = scenario_colors[
                        scenario_colors["Scenario"].str.lower() == key
                    ]
                    scenario_color_map[scen] = match["Hex Code"].iloc[0] if not match.empty else "#FFFFFF"
        else:
            for scen in scenarios:
                scenario_color_map[scen] = "#FFFFFF"

        # Fonction interne pour créer un graphique
        def create_figure(show_totals=False):
            fig, ax = plt.subplots(figsize=(16, 10), dpi=300)
            title = "Combined Scenarios Analysis"
            title += " (with Totals)" if show_totals else ""
            ax.set_title(title, fontsize=title_size, pad=80,
                         fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            ax.tick_params(axis='y', labelsize=label_size)

            for i, scenario in enumerate(scenarios):
                bottom_pos = np.zeros(len(categories))
                bottom_neg = np.zeros(len(categories))

                for contrib in [c for c in contributions_order if c in contributions]:

                    try:
                        values = (
                            percentage_table.xs((scenario, contrib), axis=1, level=[0, 1])
                            .values.flatten()
                        )
                    except KeyError:
                        continue

                    values = np.nan_to_num(values.astype(float))
                    pos_values = np.where(values > 0, values, 0)
                    neg_values = np.where(values < 0, values, 0)

                    if np.any(pos_values):
                        ax.bar(
                            x_positions + i * bar_width,
                            pos_values,
                            bar_width,
                            bottom=bottom_pos,
                            color=contrib_color_map[contrib],
                            hatch=scenario_hatch_dict.get(scenario, None)
,
                            edgecolor='black'
                        )
                        bottom_pos += pos_values

                    if np.any(neg_values):
                        ax.bar(
                            x_positions + i * bar_width,
                            neg_values,
                            bar_width,
                            bottom=bottom_neg,
                            color=contrib_color_map[contrib],
                            hatch=scenario_hatch_dict.get(scenario, None)
,
                            edgecolor='black'
                        )
                        bottom_neg += neg_values

                if show_totals:
                    scenario_clean = scenario.split("(")[-1].replace(")", "").strip()
                    try:
                        totals = total_impact_table[scenario_clean].values
                        for j, total in enumerate(totals):
                            formatted = f"{total:.2E}" if abs(total) >= 1000 else f"{total:.2f}"
                            ax.text(
                                x_positions[j] + i * bar_width,
                                bottom_pos[j] + 2,
                                formatted,
                                ha='center',
                                va='bottom',
                                rotation=90,
                                fontsize=text_size,
                                fontweight=fontweight,
                                fontstyle=fontstyle,
                                family=fontfamily
                            )
                    except KeyError:
                        pass

            ax.set_xticks(x_positions + (num_scenarios - 1) * bar_width / 2)
            ax.set_xticklabels(categories, rotation=45, ha='right',
                               fontsize=label_size, fontweight=fontweight,
                               fontstyle=fontstyle, family=fontfamily)
            ax.set_ylabel("Contribution (%)", fontsize=label_size,
                          fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            ax.axhline(0, color='black', linestyle='--', alpha=0.5)
            ax.set_ylim(top=100)
            plt.tight_layout()
            return fig

        # Graphiques principaux
        figures.append(("Main Chart", create_figure(show_totals=False)))
        figures.append(("Chart with Totals", create_figure(show_totals=True)))


        # Légende pour les contributions
        legend_fig1, ax1 = plt.subplots(figsize=(6, 2), dpi=300)
        ax1.axis('off')

        contrib_handles = [
    plt.Rectangle((0, 0), 1, 1, facecolor=color, edgecolor='black')
    for _, color in contrib_color_map.items()
]
        contrib_labels = list(contrib_color_map.keys())

        legend1 =ax1.legend(
    contrib_handles,
    contrib_labels,
    title="Contributions",
    fontsize=legend_size,
    title_fontsize=legend_size,
    ncol=2,
    loc='center'
)
        legend1.get_frame().set_linewidth(0) 
        
        figures.append(("Contributions Legend", legend_fig1))


        # Légende pour les scénarios
        legend_fig2, ax2 = plt.subplots(figsize=(6, 2), dpi=300)
        ax2.axis('off')

        scenario_handles = [
        plt.Rectangle((0, 0), 1, 1,
                  facecolor=scenario_color_map[scenario],
                  hatch=None if scenario_hatch_dict[scenario] in [None, "None"] else scenario_hatch_dict[scenario]
,
                  edgecolor='black')
        for scenario in scenarios
]
        scenario_labels = [s.split("(")[-1].replace(")", "").strip() for s in scenarios]

        legend2 = ax2.legend(
    scenario_handles,
    scenario_labels,
    title="Scenarios",
    fontsize=legend_size,
    title_fontsize=legend_size,
    ncol=2,
    loc='center'
)
        legend2.get_frame().set_linewidth(0)
        figures.append(("Scenarios Legend", legend_fig2))
        return figures

    except Exception as e:
        st.error(f"Error generating combined scenario chart: {str(e)}")
        return []
def plot_combined_graph_with_scenario_hatches_horizontal(percentage_table, total_impact_table, contributions_order):
    """
    Génère des graphiques horizontaux combinés par scénario avec hachures et légendes.
    Retourne une liste de tuples : (description, figure).
    """
    try:
        global contributions_colors, scenario_colors, font_styles, scenario_hatches
        figures = []

        fontfamily = font_styles["family"]
        fontweight = font_styles["weight"]
        fontstyle = font_styles["style"]
        title_size = font_styles["title_size"]
        label_size = font_styles["label_size"]
        legend_size = font_styles["legend_size"]
        text_size = font_styles["text_size"]

        if percentage_table.empty or total_impact_table.empty:
            st.warning("Missing data for the combined scenario chart.")
            return []

        categories = percentage_table.index
        scenarios = percentage_table.columns.levels[0]
        contributions = percentage_table.columns.levels[1]
        num_scenarios = len(scenarios)
        num_categories = len(categories)
        bar_height = min(0.9, max(0.3, 1.5 / num_scenarios))
        line_spacing = 1.5 * bar_height
        y_positions = np.arange(num_categories) * line_spacing
        fig_width = 12 + 0.5 * num_scenarios
        fig_height = max(6, num_categories * line_spacing * 0.7)
        
        

        y_positions = np.arange(len(categories)) * 1.5

        scenario_hatch_dict = {scenario: scenario_hatches.get(scenario.split("(")[-1].replace(")", "").strip().lower(), None) for scenario in scenarios}

        # Couleurs contributions
        cmap = plt.get_cmap("tab10")
        contrib_color_map = {contrib: contributions_colors.get(contrib.replace("%", "").strip().lower(), cmap(i % 10)) for i, contrib in enumerate(contributions)}

        # Couleurs scénarios
        scenario_color_map = {scen: scenario_colors.get(scen.strip().lower(), "#FFFFFF") for scen in scenarios}

        def create_figure(show_totals=False):
            fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=300)
            title = "Combined Scenarios Analysis (Horizontal)"
            title += " (with Totals)" if show_totals else ""
            ax.set_title(title, fontsize=title_size, pad=80, fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            ax.tick_params(axis='x', labelsize=label_size)

            for i, scenario in enumerate(scenarios):
                left_pos = np.zeros(len(categories))
                left_neg = np.zeros(len(categories))

                for contrib in [c for c in contributions_order if c in contributions]:
                    try:
                        values = percentage_table.xs((scenario, contrib), axis=1, level=[0, 1]).values.flatten()
                    except KeyError:
                        continue
                    values = np.nan_to_num(values.astype(float))
                    pos_values = np.where(values > 0, values, 0)
                    neg_values = np.where(values < 0, values, 0)

                    if np.any(pos_values):
                        ax.barh(y_positions + i * bar_height, pos_values, bar_height, left=left_pos, color=contrib_color_map[contrib], hatch=scenario_hatch_dict.get(scenario, None), edgecolor='black')
                        left_pos += pos_values

                    if np.any(neg_values):
                        ax.barh(y_positions + i * bar_height, neg_values, bar_height, left=left_neg, color=contrib_color_map[contrib], hatch=scenario_hatch_dict.get(scenario, None), edgecolor='black')
                        left_neg += neg_values

                if show_totals:
                    scenario_clean = scenario.split("(")[-1].replace(")", "").strip()
                    try:
                        totals = total_impact_table[scenario_clean].values
                        for j, total in enumerate(totals):
                            formatted = f"{total:.2E}" if abs(total) >= 1000 else f"{total:.2f}"
                            ax.text(left_pos[j] + 2, y_positions[j] + i * bar_height, formatted, va='center', ha='left', fontsize=text_size, fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
                    except KeyError:
                        pass

            ax.set_yticks(y_positions + (num_scenarios - 1) * bar_height / 2)
            ax.set_yticklabels(categories, fontsize=label_size, fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            ax.set_xlabel("Contribution (%)", fontsize=label_size, fontweight=fontweight, fontstyle=fontstyle, family=fontfamily)
            ax.axvline(0, color='black', linestyle='--', alpha=0.5)
            ax.set_xlim(left=min(left_neg) - 5, right=max(left_pos) + 5)
            plt.tight_layout()
            return fig

        figures.append(("Main Chart (Horizontal)", create_figure(False)))
        figures.append(("Chart with Totals (Horizontal)", create_figure(True)))

        legend_fig1, ax1 = plt.subplots(figsize=(6, 2), dpi=300)
        ax1.axis('off')
        contrib_handles = [plt.Rectangle((0, 0), 1, 1, facecolor=color, edgecolor='black') for _, color in contrib_color_map.items()]
        contrib_labels = list(contrib_color_map.keys())
        legend1 = ax1.legend(contrib_handles, contrib_labels, title="Contributions", fontsize=legend_size, title_fontsize=legend_size, ncol=2, loc='center')
        legend1.get_frame().set_linewidth(0)
        figures.append(("Contributions Legend (Horizontal)", legend_fig1))

        legend_fig2, ax2 = plt.subplots(figsize=(6, 2), dpi=300)
        ax2.axis('off')
        scenario_handles = [plt.Rectangle((0, 0), 1, 1, facecolor=scenario_color_map[scenario], hatch=None if scenario_hatch_dict[scenario] in [None, "None"] else scenario_hatch_dict[scenario], edgecolor='black') for scenario in scenarios]
        scenario_labels = [s.split("(")[-1].replace(")", "").strip() for s in scenarios]
        legend2 = ax2.legend(scenario_handles, scenario_labels, title="Scenarios", fontsize=legend_size, title_fontsize=legend_size, ncol=2, loc='center')
        legend2.get_frame().set_linewidth(0)
        figures.append(("Scenarios Legend (Horizontal)", legend_fig2))

        return figures

    except Exception as e:
        st.error(f"Error generating horizontal combined scenario chart: {str(e)}")
        return []

    




def generate_table_graph(df, table_name):
    """
    Generates a matplotlib figure displaying a DataFrame as a nicely formatted table:
    - Numeric values in scientific notation (2E)
    - First row (header) and first column highlighted with a lightblue background and bold text.
    """
    # Copy the DataFrame and format numeric values in 2E notation
    formatted_df = df.copy()
    formatted_df = formatted_df.applymap(
        lambda x: f"{x:.2E}" if isinstance(x, (int, float)) else x
    )

    num_rows, num_cols = formatted_df.shape
    fig_width = max(10, num_cols * 1.5)
    fig_height = max(6, num_rows * 0.5 + 1.5)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.set_axis_off()

    title_y_position = 0.96
    table_y_position = 0.96 - (num_rows * 0.01)
    table_y_position = max(0.70, table_y_position)

    plt.figtext(
        0.5,
        title_y_position,
        f"Table: {table_name}",
        fontsize=18,
        fontweight="bold",
        ha="center",
    )

    table = ax.table(
        cellText=formatted_df.values,
        colLabels=formatted_df.columns,
        rowLabels=formatted_df.index,
        loc="center",
        cellLoc="center",
        colWidths=[0.5] * formatted_df.shape[1],
    )

    for key, cell in table.get_celld().items():
        cell.set_fontsize(14)
        cell.set_height(0.08)

    header_color = "lightblue"
    first_col_color = "lightblue"

    # Header row
    for i in range(len(formatted_df.columns)):
        cell = table[0, i]
        cell.set_text_props(weight="bold")
        cell.set_facecolor(header_color)

    # First column
    for i in range(len(formatted_df.index)):
        cell = table[i + 1, -1]
        cell.set_text_props(weight="bold")
        cell.set_facecolor(first_col_color)

    plt.subplots_adjust(top=table_y_position, bottom=0.1)
    return fig


def find_file_path(file_name, start_directory=None):
    """
    Recursively searches for a file starting from 'start_directory'.
    If 'start_directory' is not provided, uses the user's home directory.
    """
    if start_directory is None:
        start_directory = os.path.expanduser("~")

    for root, dirs, files in os.walk(start_directory):
        if file_name in files:
            return os.path.join(root, file_name)

    return None





###############################
#           Main App
###############################

def main():
    global scenario_colors, contributions_colors, font_styles, scenario_hatches
    st.markdown(
    """
    <h1 style='text-align: center; font-size: 3em;'>
        LCA Graph Dashboard 🌍
    </h1>
    <p style='text-align: center; font-size: 1.1em;'>
        
    """,
    unsafe_allow_html=True
)

    st.markdown("""
    This tool is designed to automatically generate insightful **Life Cycle Assessment (LCA)** graphs from a structured Excel file.

    It combines a simple data entry template with a Python-powered interface to visualize your environmental impact data across multiple dimensions, including:
    - 📊 **Scenario comparison**
    - 🧩 **Relative contribution charts (vertical & horizontal)**
    - 📚 **Impact analysis by category**
    - 🎯 **Combined views of scenarios and contributions**
    """)

    st.markdown("---")

    # ----------------------------------------
    # 🔹 Étape 1 : Télécharger la template
    # ----------------------------------------
    st.header("Step 1: Prepare your data")

    st.markdown("""
    Download the Excel template below and fill in your LCA data accordingly.

    **Structuring rules:**
    - You can define multiple scenarios.
    - Each scenario can include multiple contributions.
    - A contribution can appear in several scenarios.
    - You can include as many impact categories as needed.
    
    🔎 **Note**: The Excel file already contains examples data to help you visualize the kind of charts that will be generated.
    """)

    with open("assets/LCA_template.xlsx", "rb") as f:
        excel_bytes = f.read()

    st.download_button(
        label="📥 Download Excel Template",
        data=excel_bytes,
        file_name="LCA_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("---")


    # File uploader
    st.header("Step 2: Upload your filled Excel file")
    uploaded_file = st.file_uploader("Upload your completed Excel file", type=["xlsx"], key="file_upload_main")

    if uploaded_file is not None:
        try:
            # Lire les feuilles disponibles dans le fichier
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names

            # Choix de la feuille par l'utilisateur
            selected_sheet = st.selectbox("📄 Choose the sheet to analyze", sheet_names)

            # Affichage aperçu rapide (5 premières lignes)
            preview_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5)
            with st.expander("👁️ Preview of the selected sheet"):
                st.dataframe(preview_df)
            
            # Analyze the Excel file from a specified sheet (e.g., "Feuil2")
            initial_table, combined_table, total_impact_table, scenario_names_cleaned = analyze_excel_and_generate_tables(
                uploaded_file, sheet_name=selected_sheet
            )
            # Si l'analyse a échoué
            if initial_table is None or combined_table is None or total_impact_table is None:
                st.warning("⚠️ The selected sheet does not follow the expected template. Please verify that the sheet structure matches the required format (scenarios in row 1, contributions in row 2, categories in column A starting from row 3).")
            else: 
             contributions_list = extract_contributions_from_initial_table(initial_table)
             st.sidebar.subheader("🔢 Contribution stacking order")
             contributions_order = st.sidebar.multiselect(
    "🪜 Order of contributions (bottom to top)", 
    options=contributions_list, 
    default=contributions_list,
    key="custom_contribution_order"
)
             st.write("Ordre des contributions sélectionné :", contributions_order)

 

             # Génération des couleurs par défaut
             from matplotlib import colors as mcolors


             def generate_default_colors(names, color_list):
                hex_colors = [mcolors.to_hex(c) for c in color_list]
                return {name.lower(): hex_colors[i % len(hex_colors)] for i, name in enumerate(names)}
             basic_contribution_palette = plt.get_cmap("Dark2").colors  
             
             default_scenario_colors = generate_default_colors(scenario_names_cleaned, list(mcolors.TABLEAU_COLORS.values()))
             default_contribution_colors = generate_default_colors(contributions_list, basic_contribution_palette)
             def generate_default_hatches(scenario_names):
                 hatch_options = ["//", "oo", "xx", "--", "||", "++", "..", "None"]
                 return {name.lower(): hatch_options[i % len(hatch_options)] for i, name in enumerate(scenario_names)}
             default_scenario_hatches = generate_default_hatches(scenario_names_cleaned)
             scenario_hatches = default_scenario_hatches.copy()

            # Dictionnaires globaux modifiables
             
             scenario_hatches = default_scenario_hatches.copy()

             scenario_colors = default_scenario_colors.copy()
             contributions_colors = default_contribution_colors.copy()
             font_styles = get_font_styles()

            # Interface de personnalisation dans la sidebar
             st.sidebar.subheader("🎨 Color customization")
             customize_colors = st.sidebar.checkbox("Would you like to customize the colors?")
             
             if customize_colors:
                st.sidebar.markdown("#### Scenario colors")
                for scen in scenario_names_cleaned:
                    scen_lower = scen.lower()
                    picked = st.sidebar.color_picker(f"🎯 {scen.capitalize()}", scenario_colors[scen_lower])
                    scenario_colors[scen_lower] = picked

                st.sidebar.markdown("#### Contribution colors")
                for contrib in contributions_list:
                    contrib_lower = contrib.lower()
                    picked = st.sidebar.color_picker(f"📦 {contrib}", contributions_colors[contrib_lower])
                    contributions_colors[contrib_lower] = picked
             st.sidebar.subheader("🔳 Hatch customization")
             customize_hatches = st.sidebar.checkbox("Would you like to customize the hatches?")

             if customize_hatches:
                 st.sidebar.markdown("#### Scenario hatches")
                 hatch_choices = ["///","//","/",  "xx","XX", "++", "..","O","oo", "None"]

                 for scen in scenario_names_cleaned:
                     scen_lower = scen.lower()
                     current_hatch = scenario_hatches.get(scen_lower, "None")
                     selected_hatch = st.sidebar.selectbox(
            f"Hatch for {scen.capitalize()}",
            hatch_choices,
            index=hatch_choices.index(current_hatch) if current_hatch in hatch_choices else 0,
            key=f"hatch_selector_{scen_lower}"
        )
                     scenario_hatches[scen_lower] = None if selected_hatch == "None" else selected_hatch

             if initial_table is not None:
                          scenario_tables = generate_tables_by_scenario(initial_table, scenario_names_cleaned)
                          percentage_table = generate_percentage_table(initial_table, total_impact_table)

                          st.subheader("Select Charts to Generate")
                          generate_scenario_comparison = st.checkbox("Generate Scenario Comparison Charts")
                          generate_scenario_details = st.checkbox("Generate Detailed Scenario Charts (Vertical)")
                          generate_scenario_details_horizontal = st.checkbox("Generate Detailed Scenario Charts (Horizontal)")
                          generate_category_details = st.checkbox("Generate Category-by-Category Charts")
                          combined_chart_mode = st.selectbox(
    "📊 Combined View of All Scenarios",
    ["Do not generate", "Vertical", "Horizontal"]
)
                          generate_combined_charts = combined_chart_mode == "Vertical"
                          generate_combined_charts_horizontal = combined_chart_mode == "Horizontal"

                          
                          


                # 1) Scenario Comparison
                          if generate_scenario_comparison and total_impact_table is not None:
                                st.markdown("---")
                                st.header("Scenario Comparison")
                                comparison_figures = plot_comparison_bar_chart(total_impact_table)
                                if comparison_figures:
                                     titles = ["Main Scenario Comparison", "Legend", "Scenario Comparison (with Totals)"]

                                     for i, fig in enumerate(comparison_figures):
                                            st.subheader(titles[i])
                                            st.pyplot(fig)

                                            col1, col2 = st.columns(2)

                                            # 🔹 PNG export (300 DPI)
                                            buf_png = io.BytesIO()
                                            fig.savefig(buf_png, format="png", dpi=300, bbox_inches='tight')
                                            col1.download_button(
                label="📥 Télécharger PNG (300 DPI)",
                data=buf_png.getvalue(),
                file_name=f"{titles[i].replace(' ', '_').lower()}.png",
                mime="image/png",
                key=f"png_button_{i}"
            )

                                           # 🔹 SVG export
                                            buf_svg = io.BytesIO()
                                            fig.savefig(buf_svg, format="svg", bbox_inches='tight')
                                            col2.download_button(
                label="📐 Télécharger SVG (vectoriel)",
                data=buf_svg.getvalue(),
                file_name=f"{titles[i].replace(' ', '_').lower()}.svg",
                mime="image/svg+xml",
                key=f"svg_button_{i}"
            )

                        # 2) Detailed Scenario Charts (vertical)
                          if generate_scenario_details and scenario_tables:
                               st.markdown("---")
                               st.header("Detailed Analysis by Scenario (Vertical)")
                               scenario_figures = plot_relative_contribution_by_scenario(scenario_tables, contributions_order)

                               for idx, (main_fig, legend_fig) in enumerate(scenario_figures):
                                      scenario_name = list(scenario_tables.keys())[idx]
                                      st.subheader(f"Scenario: {scenario_name}")
                                      
                                      with st.container():
                                                              
                                        col_main, col_legend = st.columns([3, 1])
                                        with col_main:
                                           st.pyplot(main_fig)

                                           col_dl1, col_dl2 = st.columns(2)

                                           # 🔽 Main figure download
                                           buf_main_png = io.BytesIO()
                                           main_fig.savefig(buf_main_png, format="png", dpi=300, bbox_inches='tight')
                                           col_dl1.download_button(
                                                   label="📥 Courbe PNG (300 DPI)",
                                                   data=buf_main_png.getvalue(),
                                                   file_name=f"{scenario_name}_vertical.png",
                                                   mime="image/png",
                                                   key=f"vertical_png_{scenario_name}"
            )

                                           buf_main_svg = io.BytesIO()
                                           main_fig.savefig(buf_main_svg, format="svg", bbox_inches='tight')
                                           col_dl2.download_button(
                label="📐 Courbe SVG",
                data=buf_main_svg.getvalue(),
                file_name=f"{scenario_name}_vertical.svg",
                mime="image/svg+xml",
                key=f"vertical_svg_{scenario_name}"
            )

                                        with col_legend:
                                          st.pyplot(legend_fig)

                                          col_leg1, col_leg2 = st.columns(2)

                                          buf_leg_png = io.BytesIO()
                                          legend_fig.savefig(buf_leg_png, format="png", dpi=300, bbox_inches='tight')
                                          col_leg1.download_button(
                label="📥 Légende PNG (300 DPI)",
                data=buf_leg_png.getvalue(),
                file_name=f"{scenario_name}_legend_vertical.png",
                mime="image/png",
                key=f"legend_vertical_png_{scenario_name}"
            )

                                        buf_leg_svg = io.BytesIO()
                                        legend_fig.savefig(buf_leg_svg, format="svg", bbox_inches='tight')
                                        col_leg2.download_button(
                label="📐 Légende SVG",
                data=buf_leg_svg.getvalue(),
                file_name=f"{scenario_name}_legend_vertical.svg",
                mime="image/svg+xml",
                key=f"legend_vertical_svg_{scenario_name}"
            )

                               st.markdown("---")

# 3) Detailed Scenario Charts (horizontal)
                          if generate_scenario_details_horizontal and scenario_tables:
                               st.markdown("---")
                               st.header("Detailed Analysis by Scenario (Horizontal)")
                               horizontal_figures = plot_relative_contribution_by_scenario_horizontal(scenario_tables, contributions_order)

                               for idx, (main_fig, legend_fig) in enumerate(horizontal_figures):
                                      scenario_name = list(scenario_tables.keys())[idx]
                                      st.subheader(f"Scenario: {scenario_name} (Horizontal)")
                                      
                                      with st.container():
                                                              
                                        col_main, col_legend = st.columns([3, 1])
                                        with col_main:
                                           st.pyplot(main_fig)

                                           col_dl1, col_dl2 = st.columns(2)

                                           # 🔽 Main figure download
                                           buf_main_png = io.BytesIO()
                                           main_fig.savefig(buf_main_png, format="png", dpi=300, bbox_inches='tight')
                                           col_dl1.download_button(
                                                   label="📥 Courbe PNG (300 DPI)",
                                                   data=buf_main_png.getvalue(),
                                                   file_name=f"{scenario_name}_horizontal.png",
                                                   mime="image/png",
                                                   key=f"horizontal_png_{scenario_name}"
            )

                                           buf_main_svg = io.BytesIO()
                                           main_fig.savefig(buf_main_svg, format="svg", bbox_inches='tight')
                                           col_dl2.download_button(
                label="📐 Courbe SVG",
                data=buf_main_svg.getvalue(),
                file_name=f"{scenario_name}_horizontal.svg",
                mime="image/svg+xml",
                key=f"horizontal_svg_{scenario_name}"
            )

                                        with col_legend:
                                          st.pyplot(legend_fig)

                                          col_leg1, col_leg2 = st.columns(2)

                                          buf_leg_png = io.BytesIO()
                                          legend_fig.savefig(buf_leg_png, format="png", dpi=300, bbox_inches='tight')
                                          col_leg1.download_button(
                label="📥 Légende PNG (300 DPI)",
                data=buf_leg_png.getvalue(),
                file_name=f"{scenario_name}_legend_horizontal.png",
                mime="image/png",
                key=f"legend_horizontal_png_{scenario_name}"
            )

                                        buf_leg_svg = io.BytesIO()
                                        legend_fig.savefig(buf_leg_svg, format="svg", bbox_inches='tight')
                                        col_leg2.download_button(
                label="📐 Légende SVG",
                data=buf_leg_svg.getvalue(),
                file_name=f"{scenario_name}_legend_horizontal.svg",
                mime="image/svg+xml",
                key=f"legend_horizontal_svg_{scenario_name}"
            )

                               st.markdown("---")

                # 4) Category-by-Category Charts
                          if generate_category_details and percentage_table is not None and total_impact_table is not None:
                                              st.markdown("---")
                                              st.header("Detailed Analysis by Impact Category")
                                              category_figures = plot_stacked_bar_by_category(initial_table, total_impact_table, contributions_order)

                                              if category_figures:
                                                  for main_fig, legend_fig, category_name in category_figures:
                                                      st.subheader(f"Category: {category_name}")
                          
                                                      col1, col2 = st.columns([4, 1])
                                                      with col1:
                                                          st.pyplot(main_fig)
                                                      with col2:
                                                          st.pyplot(legend_fig)
                          
                                                      plt.close(main_fig)
                                                      plt.close(legend_fig)

                                                      st.markdown("---")

                          # 5) Combined Charts
                          if generate_combined_charts and percentage_table is not None and total_impact_table is not None:
                              st.markdown("---")
                              st.header("Combined View of All Scenarios")
                              combined_figures = plot_combined_graph_with_scenario_hatches(percentage_table, total_impact_table, contributions_order)

                          
                              if combined_figures:
                                  # Séparer légendes et courbes
                                  legends = {}
                                  charts = []
                          
                                  for name, fig in combined_figures:
                                      if "Legend" in name:
                                          legends[name] = fig
                                      else:
                                          charts.append((name, fig))
                          
                                  # 📌 Afficher d’abord les légendes avec téléchargement
                                  for legend_name, legend_fig in legends.items():
                                      st.subheader(legend_name)
                                      st.pyplot(legend_fig)
                          
                                      col1, col2 = st.columns(2)
                          
                                      buf_png = io.BytesIO()
                                      legend_fig.savefig(buf_png, format="png", dpi=300, bbox_inches='tight')
                                      col1.download_button(
                                          label="📥 Légende PNG (300 DPI)",
                                          data=buf_png.getvalue(),
                                          file_name=f"{legend_name.replace(' ', '_').lower()}.png",
                                          mime="image/png",
                                          key=f"legend_png_{legend_name}"
                                      )

                                      buf_svg = io.BytesIO()
                                      legend_fig.savefig(buf_svg, format="svg", bbox_inches='tight')
                                      col2.download_button(
                                          label="📐 Légende SVG",
                                          data=buf_svg.getvalue(),
                                          file_name=f"{legend_name.replace(' ', '_').lower()}.svg",
                                          mime="image/svg+xml",
                                          key=f"legend_svg_{legend_name}"
                                      )
                          
                                      st.markdown("---")
                          
                                  # 🖼️ Puis afficher chaque graphique principal avec téléchargement
                                  for name, fig in charts:
                                      st.subheader(name)
                                      st.pyplot(fig)
                          
                                      col1, col2 = st.columns(2)
                          
                                      buf_png = io.BytesIO()
                                      fig.savefig(buf_png, format="png", dpi=300, bbox_inches='tight')
                                      col1.download_button(
                                          label="📥 PNG (300 DPI)",
                                          data=buf_png.getvalue(),
                                          file_name=f"{name.replace(' ', '_').lower()}.png",
                                          mime="image/png",
                                          key=f"combined_png_{name}"
                                      )
                          
                                      buf_svg = io.BytesIO()
                                      fig.savefig(buf_svg, format="svg", bbox_inches='tight')
                                      col2.download_button(
                                          label="📐 SVG",
                                          data=buf_svg.getvalue(),
                                          file_name=f"{name.replace(' ', '_').lower()}.svg",
                                          mime="image/svg+xml",
                                          key=f"combined_svg_{name}"
                                      )
                          
                                      st.markdown("---")
                        # 6) Combined Charts - Horizontal
                          if generate_combined_charts_horizontal and percentage_table is not None and total_impact_table is not None:
                              st.markdown("---")
                              st.header("Combined View of All Scenarios (Horizontal)")
                              combined_horizontal_figures = plot_combined_graph_with_scenario_hatches_horizontal(percentage_table, total_impact_table, contributions_order)
                              if combined_horizontal_figures:
                                  legends_h = {}
                                  charts_h = []
                                  for name, fig in combined_horizontal_figures:
                                      if "Legend" in name:
                                          legends_h[name] = fig
                                      else:
                                          charts_h.append((name, fig))
                                  for legend_name, legend_fig in legends_h.items():
                                      st.subheader(legend_name)
                                      st.pyplot(legend_fig)
                                      col1, col2 = st.columns(2)
                                      buf_png = io.BytesIO()
                                      legend_fig.savefig(buf_png, format="png", dpi=300, bbox_inches='tight')
                                      col1.download_button(
                label="📥 Légende PNG (300 DPI)",
                data=buf_png.getvalue(),
                file_name=f"{legend_name.replace(' ', '_').lower()}.png",
                mime="image/png",
                key=f"legend_horizontal_png_{legend_name}"
            )
                                      buf_svg = io.BytesIO()
                                      legend_fig.savefig(buf_svg, format="svg", bbox_inches='tight')
                                      col2.download_button(
                label="📐 Légende SVG",
                data=buf_svg.getvalue(),
                file_name=f"{legend_name.replace(' ', '_').lower()}.svg",
                mime="image/svg+xml",
                key=f"legend_horizontal_svg_{legend_name}"
            )
                                      st.markdown("---")
                              for name, fig in charts_h:
                                  st.subheader(name)
                                  st.pyplot(fig)
                                  col1, col2 = st.columns(2)
                                  buf_png = io.BytesIO()
                                  fig.savefig(buf_png, format="png", dpi=300, bbox_inches='tight')
                                  col1.download_button(
                label="📥 PNG (300 DPI)",
                data=buf_png.getvalue(),
                file_name=f"{name.replace(' ', '_').lower()}.png",
                mime="image/png",
                key=f"combined_horizontal_png_{name}"
            )
                                  buf_svg = io.BytesIO()
                                  fig.savefig(buf_svg, format="svg", bbox_inches='tight')
                                  col2.download_button(
                label="📐 SVG",
                data=buf_svg.getvalue(),
                file_name=f"{name.replace(' ', '_').lower()}.svg",
                mime="image/svg+xml",
                key=f"combined_horizontal_svg_{name}"
            )
                                  st.markdown("---")
                                  
                                  
        except Exception as e:
            st.error(f"Error while processing the file: {str(e)}")


if __name__ == "__main__":
    main()