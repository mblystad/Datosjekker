import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk  # Import for å vise logoen
import os
import openpyxl

# Funksjon for å sammenligne dato-kolonnene fra flere Excel-filer
def compare_dates(files):
    try:
        # Leser inn Excel-filene i DataFrames
        dataframes = [pd.read_excel(file, index_col=0) for file in files]

        # Velger "Dato"-kolonnene fra hver DataFrame
        dates = [df["Dato"] for df in dataframes]

        # Rapportfil for lagring av resultatene
        report_file = "rapport_over_dato_kraesjer.txt"
        with open(report_file, 'w') as f:
            if len(files) == 2:
                # Finn overlappende datoer mellom to filer
                overlapping_dates = dates[0][dates[0].isin(dates[1])].dropna()

                if not overlapping_dates.empty:
                    f.write(
                        f"Overlappende datoer mellom {os.path.basename(files[0])} og {os.path.basename(files[1])}:\n")
                    for i, date in enumerate(overlapping_dates):
                        f.write(f"Rad {i + 1}: {date}\n")
                else:
                    f.write(
                        f"Ingen overlappende datoer mellom {os.path.basename(files[0])} og {os.path.basename(files[1])}.\n")
            else:
                # Finn overlappende datoer mellom flere filer
                for i, date1 in enumerate(dates):
                    for j, date2 in enumerate(dates[i + 1:], start=i + 1):
                        overlapping_dates = date1[date1.isin(date2)].dropna()
                        if not overlapping_dates.empty:
                            f.write(
                                f"Overlappende datoer mellom {os.path.basename(files[i])} og {os.path.basename(files[j])}:\n")
                            for k, date in enumerate(overlapping_dates):
                                f.write(f"Rad {k + 1}: {date}\n")
                            f.write("\n")
                        else:
                            f.write(
                                f"Ingen overlappende datoer mellom {os.path.basename(files[i])} og {os.path.basename(files[j])}.\n\n")

        # Meld bruker om at rapporten er lagret
        messagebox.showinfo("Resultat", f"Overlappende datoer er lagret i rapporten: {os.path.abspath(report_file)}")

    except Exception as e:
        messagebox.showerror("Feil", f"Det oppsto en feil:\n{str(e)}")


# Funksjon for å velge filer og vise filstien i tilhørende felt
def browse_files(file_var):
    file_path = filedialog.askopenfilename(filetypes=[("Excel-filer", "*.xlsx")])
    if file_path:  # Hvis filen er valgt
        file_var.set(file_path)


# Funksjon for å starte sammenligning
def start_comparison():
    files = [var.get() for var in file_paths if var.get()]
    if len(files) > 1:
        compare_dates(files)
    else:
        messagebox.showwarning("Input påkrevd", "Vennligst velg minst to Excel-filer for å sammenligne.")


# Oppsett av tkinter GUI
root = tk.Tk()
root.title("CYDBYWYD")
root.configure(bg='#9faba7')  # Setter bakgrunnsfarge for hele vinduet

# Legger til logoen i GUI
try:
    logo_img = Image.open("logo.png")
    logo_img = logo_img.resize((180, 150))  # Endre størrelsen på bildet for å passe
    logo = ImageTk.PhotoImage(logo_img)
    tk.Label(root, image=logo, bg='#9faba7').grid(row=0, column=0, columnspan=3, pady=10)
except Exception as e:
    messagebox.showwarning("Advarsel", f"Logoen kunne ikke lastes inn: {str(e)}")

# Variabler for å lagre filstiene
file_paths = []

# GUI-oppsett for flere filvalg
for i in range(3):  # Endre 3 til ønsket antall filer
    file_var = tk.StringVar()
    file_paths.append(file_var)

    tk.Label(root, text=f"Velg Excel-fil {i + 1}:", bg='#9faba7').grid(row=i + 1, column=0, padx=10, pady=10)
    tk.Entry(root, textvariable=file_var, width=50).grid(row=i + 1, column=1, padx=10, pady=10)
    tk.Button(root, text="Bla gjennom", command=lambda fv=file_var: browse_files(fv), bg='#9faba7').grid(row=i + 1,
                                                                                                         column=2,
                                                                                                         padx=10,
                                                                                                         pady=10)

tk.Button(root, text="Sammenlign Datoer", command=start_comparison, width=20, bg='#9faba7').grid(row=4, column=1,
                                                                                                 pady=20)

# Kjører GUI-løkken
root.mainloop()
