"""
Tkinter gUI application to process Excel files
"""

import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

TEMP_FILE = "temp.xlsx"
HEADERS = [
    "TYPE",
    "Raison C/F",
    "Date",
    "Lot",
    "Désignation",
    "Poids",
    "UN",
    "PU",
    "Résultat",
]
YELLOW = "FFFF00"
RED = "FF0000"
LIGHTBLUE = "A0D0FF"
GRAY = "808080"
LIGHTGRAY = "D0D0D0"
WHITE = "FFFFFF"
BLACK = "000000"

output_dir = ""


def process_excel_file():
    """
    Function called when the "Open" button is clicked.
    """
    excel_file = open_file()
    df = clean_dataframe(excel_file)
    buyer_groups = process_dataframe(df)
    wb = create_excel_file(buyer_groups)
    wb = apply_styles(wb)
    ddmmyy = datetime.now().strftime("%d%m%y")
    final_file_name = f"MARGES PAR ARRIVAGES {ddmmyy}.xlsx"
    if not os.path.exists(output_dir):
        show_warning("Veuillez sélectionner un répertoire valide.")
        return
    else:
        wb.save(os.path.join(output_dir, final_file_name))
    print("Saved.")


def apply_styles(wb):
    """
    Read a temp keyword in each row to apply specific styles.
    Delete temp columns
    wb : openpyxl Workbook
    """
    ws = wb.active
    thin = Side(border_style="thin", color="000000")
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "TYPE":
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=GRAY)
                    cell.font = Font(bold=True, color=WHITE)
            if cell.value == "VENTE":
                for cell in row:
                    cell.font = Font(color=RED)
            if cell.value == "Corbeille":
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=LIGHTBLUE)
            if cell.value == "negative":
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=YELLOW)
                    cell.font = Font(color="FF0000", bold=True)
            if cell.value == "neg_result":
                row[8].fill = PatternFill("solid", fgColor=YELLOW)
                row[8].font = Font(color=RED, bold=True)
            if cell.value == "Sous-total":
                for cell in row:
                    cell.fill = PatternFill(fgColor=LIGHTGRAY)
            if cell.value == "Total":
                for cell in row:
                    cell.fill = PatternFill(fgColor=LIGHTGRAY)
                    cell.font = Font(bold=True, color=BLACK)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    # Delete useless last columns
    ws.delete_cols(10, amount=4)
    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Lettre de la colonne (A, B, C, ...)

        # Calculer la largeur maximale dans chaque colonne
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        # Ajuster la largeur de la colonne
        adjusted_width = (max_length + 2) * 1.2  # Ajout de marge pour plus d'esthétique
        ws.column_dimensions[col_letter].width = adjusted_width
    # Insert date before the first row
    ws.insert_rows(1)
    ws["A1"] = f"ARRIVAGES DU {datetime.now().strftime("%d/%m/%y")}"
    ws["A1"].font = Font(bold=True, color=RED)
    ws["A1"].alignment = Alignment(horizontal="left")
    return wb


def create_excel_file(buyer_groups):
    wb = Workbook()
    ws = wb.active
    group_total = 0
    for group in buyer_groups:
        ws.append(HEADERS)
        for lot in group:
            group_total += lot["result_total"]
            for r in dataframe_to_rows(lot["df"], index=False, header=False):
                if lot["result_total"] < 0:
                    r.append("negative")
                if r[0] == "TYPE":
                    r.append("header")
                if r[1] == "Corbeille":
                    r.append("recyclebin")
                ws.append(r)
            if lot["result_total"] >= 0.0:
                ws.append(
                    [
                        "Sous-total",
                        "",
                        "",
                        "",
                        "",
                        lot["weight_total"],
                        "",
                        "",
                        lot["result_total"],
                    ]
                )
            else:
                ws.append(
                    [
                        "Sous-total",
                        "",
                        "",
                        "",
                        "",
                        lot["weight_total"],
                        "",
                        "",
                        lot["result_total"],
                        "neg_result",
                    ]
                )
        ws.append(["Total", "", "", "", "", "", "", "", group_total])
        ws.append([])

    return wb


def create_excel_styles():
    return {
        "header": NamedStyle(
            name="header",
            font=Font(bold=True, color="FFFFFF"),
            fill=PatternFill("solid", fgColor="505050"),
        ),
        "sell": NamedStyle(name="sell", font=Font(color="DD0000")),
        "recyclebin": NamedStyle(fill=PatternFill("solid", fgColor="80D0FF")),
        "subtotal": NamedStyle(
            name="subtotal",
            font=Font(bold=True, color="000000"),
            fill=PatternFill("solid", fgColor="A0A0A0"),
        ),
        "neg_result": NamedStyle(
            name="neg_result",
            font=Font(bold=True),
            fill=PatternFill("solid", fgColor="FFc520"),
        ),
        "neg_result_lines": NamedStyle(
            name="neg_result_lines", fill=PatternFill("solid", fgColor="FFc520")
        ),
        "yellow_fill": PatternFill("solid", fgColor="FFFF00"),
        "blue_fill": PatternFill("solid", fgColor="4080FF"),
        "red_font": Font(color="FF0000"),
    }


def process_dataframe(df):
    # Add a temporary column to store row formats
    df["row_format"] = ""
    lot_df = divide_by_lot(df)
    lots = [calculate_subtotals(lot_df) for lot_df in lot_df]
    # Group lots by sellers ("Raison C/F"/"TYPE")
    # "TYPE" must be "ACHAT"
    buyers = group_by_buyers(lots)
    return buyers


def group_by_buyers(lots):
    """
    Each lot has one seller ("Raison C/F"/"TYPE")
    Groupe lots by sellers and append lots list to buyer_groups list
    """
    buyer_groups = []
    current_buyer = lots[0]["df"]["Raison C/F"][0]
    current_buyer_group = []
    for lot in lots:
        if lot["df"]["Raison C/F"][0] == current_buyer:
            current_buyer_group.append(lot)
        else:
            current_buyer = lot["df"]["Raison C/F"][0]
            buyer_groups.append(current_buyer_group)
            current_buyer_group = [lot]

    return buyer_groups


def divide_by_lot(df):
    """
    Returns a list of DataFrames, one for each lot.
    Only the first 11 characters of the lot number are considered.
    """
    lot_dataframes = []
    # Only the first 11 characters of the lot number are considered
    current_lot_number = str(df["Lot"][0][:11])
    # Initialize the current lot dataframe with the same columns as the original dataframe
    current_lot_dataframe = pd.DataFrame(columns=df.columns)
    # current_lot_dataframe = pd.concat(
    #     [current_lot_dataframe, df.iloc[0]], ignore_index=True
    # )
    for _, row in df.iterrows():
        if row["TYPE"] == "ACHAT" or row["TYPE"] == "VENTE":
            if str(row["Lot"])[:11] == current_lot_number:
                current_lot_dataframe = pd.concat(
                    [current_lot_dataframe, row.to_frame().T], ignore_index=True
                )
            else:
                lot_dataframes.append(current_lot_dataframe)
                current_lot_dataframe = pd.DataFrame(columns=df.columns)
                current_lot_dataframe = pd.concat(
                    [current_lot_dataframe, row.to_frame().T], ignore_index=True
                )
                current_lot_number = str(row["Lot"])[:11]
    lot_dataframes.append(current_lot_dataframe)
    return lot_dataframes


def calculate_subtotals(lot_df: pd.DataFrame):
    """
    Add a new row after each lot with the subtotal of "Poids" and "Résultat"
    """
    lot_df["Poids"] = lot_df["Poids"].astype(float)
    lot_df["Résultat"] = lot_df["Résultat"].astype(float)
    weight_total = round(lot_df["Poids"].sum(), 2)
    result_total = round(lot_df["Résultat"].sum(), 2)
    lot = {
        "df": lot_df,
        "weight_total": weight_total,
        "result_total": result_total,
        "negative": result_total < 0,
    }
    return lot


def clean_dataframe(excel_file: str):
    """Remove duplicates and unwanted columns/rows from the DataFrame."""
    df = pd.read_excel(excel_file)
    df.to_csv("temp1.csv", index=False)
    df.drop_duplicates(inplace=True)
    df.sort_values(by=["Lot", "TYPE"], inplace=True)
    df = df[df["Code C/F"] != "-REGUL"]
    df.drop(
        columns=["Code C/F", "Commande", "Article", "Colis", "Pièces"], inplace=True
    )
    # Remove index from dataframe
    df.reset_index(drop=True, inplace=True)
    df.to_csv("temp2.csv", index=False)
    if df.empty:
        raise ValueError("Le fichier Excel est vide")
    return df


def open_file() -> str:
    try:
        file_path = filedialog.askopenfilename(
            filetypes=[("Fichier Excel", "*.xlsx *.xls")],
            title="Sélectionnez un fichier",
        )
        if not file_path:
            print("Aucun fichier sélectionné")
        wb = load_workbook(file_path)
        wb.save(TEMP_FILE)
        return TEMP_FILE
    except Exception as e:
        raise ValueError(
            f"Une erreur est survenue lors de la sélection du fichier: {e}"
        )


def select_output_directory():
    """
    Permet à l'utilisateur de sélectionner un répertoire pour le fichier final.
    """
    global output_dir
    try:
        selected_dir = filedialog.askdirectory(title="Sélectionnez un répertoire")
        if not selected_dir:
            show_warning("Aucun répertoire sélectionné.")
            return
        output_dir = selected_dir
        directory_label.config(text=f"Répertoire sélectionné : {output_dir}")
        return output_dir
    except Exception as e:
        raise ValueError(
            f"Une erreur est survenue lors de la sélection du répertoire: {e}"
        )


def run():
    root = tk.Tk()
    root.title("Saprimex")
    root.geometry("400x300")
    root.resizable(False, False)

    open_button = tk.Button(root, text="Ouvrir", command=process_excel_file)
    open_button.pack(pady=20)

    output_dir_button = tk.Button(
        root, text="Sélectionner un répertoire", command=select_output_directory
    )
    output_dir_button.pack(pady=20)

    global directory_label
    directory_label = tk.Label(root, text=f"Répertoire sélectionné : {output_dir}")
    directory_label.pack(pady=20)

    root.mainloop()


def show_warning(msg):
    messagebox.showwarning("Attention", msg)


if __name__ == "__main__":
    run()
