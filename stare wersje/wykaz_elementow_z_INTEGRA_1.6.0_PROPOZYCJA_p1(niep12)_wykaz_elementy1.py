import pandas as pd
import re

# Wczytaj dane
wykaz_df = pd.read_excel("wykaz.xlsx", dtype=str)
elementy_df = pd.read_excel("elementy1.xlsx", dtype=str)

# Przygotowanie kolumny wynikowej
wykaz_df['propozycja'] = ""

# Uzupełnij puste pola
elementy_df['Referencja'] = elementy_df['Referencja'].fillna('')

# Iteracja po wykazie
for idx, row in wykaz_df.iterrows():
    nazwa = row.get('Nazwa')
    if not isinstance(nazwa, str) or nazwa.strip() == "":
        continue
    nazwa = nazwa.strip()

    # Tworzymy regex wzorzec np. dla "76057949_p1" będzie to '76057949_p1(?!\d)'
    # oznacza: dopasuj tylko jeśli po p1 nie ma cyfry (czyli nie p12, p13 itd.)
    regex = re.escape(nazwa) + r'(?!\d)'

    # Szukamy dopasowań
    dopasowania = elementy_df[elementy_df['Referencja'].str.contains(regex, na=False, regex=True)]

    if not dopasowania.empty:
        match = dopasowania.iloc[0]
        wykaz_df.at[idx, 'propozycja'] = match.get('Referencja1', "")

# Zapisz wynik
wykaz_df.to_excel("propozycje.xlsx", index=False)
print("Dopasowanie zakończone. Wyniki zapisano w 'propozycje.xlsx'.")
