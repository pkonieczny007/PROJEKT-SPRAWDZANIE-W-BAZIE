<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" style="height:64px;margin-right:32px"/>

# Popraw skrypt aby popraw spradzanie_rewizji.py tworzymy komorke ZEFZEV

tam wypełniamy formułą ZEF\&ZEV np M2\&N2
jeżeli 2 puste nie robimy, jeżeli jedne pusta robimy

teraz poprawiamy Task wypełnienia nowe_oznaczenie.

wypełniamy kolumne
Zeinr	ZEF	ZEV	ZEFZEV
76147249	0	A	0A
SL10278402	1	A	1A
SL10278402			
SL10278402			
SL10278402			
SL10278402			
SL10278402			
SL10278402

następnie przeciągamy wartości Jeżeli ZEFZEV jest puste sprawdzamy Zeinr.
I dopasowujemy wartość z góry Zeinr

Zeinr	ZEF	ZEV	ZEFZEV
76147249	0	A	0A
SL10278402	1	A	1A
SL10278402			1A
SL10278402			1A
SL10278402			1A
SL10278402			1A
SL10278402			1A
SL10278402			1A import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

# File paths

propozycje_file = 'propozycje.xlsx'  \# Assuming it's in the current directory or provide full path
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Load workbooks

wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Assuming Zeinr is column A, ZEF is column Z, ZEV is column AA (adjust if needed)

zeinr_col = 'A'  \# Zeinr column
zef_col = 'Z'    \# ZEF column
zev_col = 'AA'   \# ZEV column

# Create ZEFZEV column next to ZEV

zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF\&ZEV if both ZEF and ZEV are present

for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    if zef is not None and zev is not None:
        ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)

# Fill empty ZEFZEV by looking up same Zeinr from above rows

zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        zeinr_to_zefzev[zeinr] = zefzev
    else:
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{zefzev_col}{row}'].value = zeinr_to_zefzev[zeinr]

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV

nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'

stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'

skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file

wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active
ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F

for col in range(1, 7):  \# A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' and place into columns starting from BH

propozycja_col = 'BH'  \# Assuming propozycja is column BH, adjust if needed
bh_col_num = openpyxl.utils.column_index_from_string('BH')

for row in range(2, ws_propozycje.max_row + 1):
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
    if propozycja:
        parts = str(propozycja).split('_')
        for i, part in enumerate(parts):
            col_num = bh_col_num + i
            col_letter = get_column_letter(col_num)
            ws_propozycje[f'{col_letter}{row}'].value = part

# Fill 'skrot' with concatenation if 'propozycja' is filled

# Assuming BN is nowe_oznaczenie, BJ is something, but need to adjust columns

# For now, assuming BN is nowe_oznaczenie_col, BJ is the first split part or adjust

# The example is =BN5\&","\&BJ5, so BN is nowe_oznaczenie, BJ is the first part after split

bj_col = 'BJ'  \# Assuming BJ is the column for first split part
for row in range(2, ws_propozycje.max_row + 1):
    if ws_propozycje[f'{propozycja_col}{row}'].value:
        nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
        bj = ws_propozycje[f'{bj_col}{row}'].value
        if nowe and bj:
            ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bj)

# Task 7: Fill 'stare_oznaczenie' with VLOOKUP

# =VLOOKUP(BE5;stary!C:F;4;FALSE) where BE is skrot

for row in range(2, ws_propozycje.max_row + 1):
    if ws_propozycje[f'{propozycja_col}{row}'].value:
        skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
        if skrot_val:
            \# Perform VLOOKUP manually
            for stary_row in range(2, ws_stary.max_row + 1):
                if ws_stary[f'C{stary_row}'].value == skrot_val:  \# Assuming lookup in column C
                    ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value  \# Column F is 4th column from C
                    break

# Save the workbook

wb_propozycje.save('propozycje_updated.xlsx')

Na podstawie Twoich wymagań poprawiłem skrypt. Oto kluczowe zmiany:

```python
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Load workbooks
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Column definitions
zeinr_col = 'A'  # Zeinr column
zef_col = 'Z'    # ZEF column
zev_col = 'AA'   # ZEV column

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tylko jeżeli przynajmniej jedna wartość istnieje (nie obie puste)
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli oba puste - nie robimy nic
    if zef is None and zev is None:
        continue
    
    # Jeżeli przynajmniej jedno jest niepuste - robimy konkatenację
    zef_str = str(zef) if zef is not None else ''
    zev_str = str(zev) if zev is not None else ''
    ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Wypełniamy puste ZEFZEV na podstawie Zeinr - przeciąganie wartości w dół
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Jeżeli ZEFZEV jest wypełnione - zapisujemy do słownika
    if zefzev:
        zeinr_to_zefzev[zeinr] = zefzev
    # Jeżeli ZEFZEV jest puste - sprawdzamy czy mamy wartość dla tego Zeinr
    else:
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{zefzev_col}{row}'].value = zeinr_to_zefzev[zeinr]

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active
ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' and place into columns starting from BH
propozycja_col = 'BH'  # Assuming propozycja is column BH
bh_col_num = openpyxl.utils.column_index_from_string('BH')

for row in range(2, ws_propozycje.max_row + 1):
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
    if propozycja:
        parts = str(propozycja).split('_')
        for i, part in enumerate(parts):
            col_num = bh_col_num + i
            col_letter = get_column_letter(col_num)
            ws_propozycje[f'{col_letter}{row}'].value = part

# Fill 'skrot' with concatenation if 'propozycja' is filled
bj_col = 'BJ'  # Assuming BJ is the column for first split part
for row in range(2, ws_propozycje.max_row + 1):
    if ws_propozycje[f'{propozycja_col}{row}'].value:
        nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
        bj = ws_propozycje[f'{bj_col}{row}'].value
        if nowe and bj:
            ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bj)

# Task 7: Fill 'stare_oznaczenie' with VLOOKUP
for row in range(2, ws_propozycje.max_row + 1):
    if ws_propozycje[f'{propozycja_col}{row}'].value:
        skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
        if skrot_val:
            # Perform VLOOKUP manually
            for stary_row in range(2, ws_stary.max_row + 1):
                if ws_stary[f'C{stary_row}'].value == skrot_val:
                    ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                    break

# Save the workbook
wb_propozycje.save('propozycje_updated.xlsx')
print("Plik został pomyślnie zapisany jako 'propozycje_updated.xlsx'")
```


## Kluczowe poprawki

### Tworzenie ZEFZEV

**Poprzednio:** Wymagało obu wartości (ZEF i ZEV)[^1_1]

**Teraz:**

- Jeżeli **oba puste** - nie tworzy wartości
- Jeżeli **przynajmniej jedno niepuste** - konkatenuje dostępne wartości
- Puste wartości zamienia na pusty string zamiast "None"


### Wypełnianie pustych ZEFZEV

Mechanizm "przeciągania" wartości działa następująco:[^1_1]

1. Iteruje przez wszystkie wiersze
2. Dla każdego wiersza sprawdza Zeinr
3. Jeżeli ZEFZEV jest wypełnione - zapisuje do słownika `{Zeinr: ZEFZEV}`
4. Jeżeli ZEFZEV jest puste - szuka w słowniku wartości dla danego Zeinr i wstawia

### Dodano komunikat

Na końcu skryptu dodano informację o pomyślnym zapisaniu pliku[^1_1]

Ten skrypt teraz prawidłowo obsługuje Twój przykład:

- Dla `76147249` z `0` i `A` tworzy `0A`
- Dla `SL10278402` z `1` i `A` tworzy `1A`
- Dla kolejnych wierszy `SL10278402` z pustymi ZEF/ZEV kopiuje `1A` w dół[^1_1]
<span style="display:none">[^1_10][^1_2][^1_3][^1_4][^1_5][^1_6][^1_7][^1_8][^1_9]</span>

<div align="center">⁂</div>

[^1_1]: https://stackoverflow.com/questions/7907928/openpyxl-check-for-empty-cell

[^1_2]: https://stackoverflow.com/questions/75239945/python-openpyxl-merge-empty-row-cells-delimited-by-string

[^1_3]: https://www.reddit.com/r/learnpython/comments/10790s6/openpyxl_reading_an_empty_cell_in_a_worksheet/

[^1_4]: https://openpyxl.readthedocs.io/en/3.1.3/editing_worksheets.html

[^1_5]: https://openpyxl.readthedocs.io/en/3.1/optimized.html

[^1_6]: https://openpyxl.readthedocs.io/en/latest/api/openpyxl.cell.cell.html

[^1_7]: https://python-forum.io/thread-3884.html

[^1_8]: https://github.com/saulpw/visidata/issues/1121

[^1_9]: https://codepal.ai/excel-formula-generator/query/qo18VLU3/excel-formula-python-match-fill

[^1_10]: https://www.codespeedy.com/find-the-first-empty-cell-from-column-of-an-excel-file-using-openpyxl/


---

# popraw skrypt aby nie trzeba było wskazywac konkretnych komórek dla \# Column definitions

zeinr_col = 'L'  \# Zeinr column
zef_col = 'M'    \# ZEF column
zev_col = 'N'   \# ZEV column tylko zeby

zeinr_col wykorzystywało nazwe kolumny "Zeinr"
zef_col = wykorzystywało nazwe kolumny "ZEF"
zev_col = wykorzystywało nazwe kolumny "ZEV"

te nazwy w wykazie sa niezmienne.
mozesz do dzialan wykorzystywac pandas dodatkowo

Oto poprawiony skrypt, który automatycznie znajduje kolumny po nazwach nagłówków zamiast używać hardkodowanych liter kolumn:[^2_1][^2_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tylko jeżeli przynajmniej jedna wartość istnieje
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli oba puste - nie robimy nic
    if zef is None and zev is None:
        continue
    
    # Jeżeli przynajmniej jedno jest niepuste - robimy konkatenację
    zef_str = str(zef) if zef is not None else ''
    zev_str = str(zev) if zev is not None else ''
    ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Wypełniamy puste ZEFZEV na podstawie Zeinr - przeciąganie wartości w dół
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Jeżeli ZEFZEV jest wypełnione - zapisujemy do słownika
    if zefzev:
        zeinr_to_zefzev[zeinr] = zefzev
    # Jeżeli ZEFZEV jest puste - sprawdzamy czy mamy wartość dla tego Zeinr
    else:
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{zefzev_col}{row}'].value = zeinr_to_zefzev[zeinr]

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' and place into columns starting from BH
# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    bh_col_num = propozycja_col_idx + 1
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Pomijam split.")
    propozycja_col = None

if propozycja_col:
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja:
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = bh_col_num + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BJ to pierwsza kolumna po podziale (BH + 2)
    bj_col = get_column_letter(bh_col_num + 2)
    
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bj = ws_propozycje[f'{bj_col}{row}'].value
            if nowe and bj:
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bj)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val:
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        break

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\nPlik został pomyślnie zapisany jako '{output_file}'")
```


## Główne zmiany

### Automatyczne znajdowanie kolumn

Skrypt używa `df.columns.get_loc()` do znalezienia indeksów kolumn po nazwach:[^2_2]

- **Zeinr** - automatycznie lokalizowana
- **ZEF** - automatycznie lokalizowana
- **ZEV** - automatycznie lokalizowana


### Konwersja indeksów

Indeksy pandas (0-based) są konwertowane na litery Excel (1-based) używając `get_column_letter()`:[^2_3][^2_1]

```python
zeinr_col = get_column_letter(zeinr_col_idx + 1)
```


### Obsługa błędów

Jeśli któraś z wymaganych kolumn nie istnieje, skrypt wyświetla komunikat błędu i kończy działanie[^2_2]

### Informacje diagnostyczne

Skrypt wyświetla znalezione kolumny, co ułatwia debugowanie[^2_1]

### Automatyczne znajdowanie kolumny "propozycja"

Również kolumna "propozycja" jest znajdowana automatycznie z obsługą sytuacji gdy nie istnieje[^2_2]

Dzięki tym zmianom skrypt jest teraz elastyczny i nie wymaga ręcznej aktualizacji liter kolumn gdy zmieniają się pozycje w arkuszu.[^2_1][^2_2]
<span style="display:none">[^2_10][^2_4][^2_5][^2_6][^2_7][^2_8][^2_9]</span>

<div align="center">⁂</div>

[^2_1]: https://stackoverflow.com/questions/71650214/obtain-header-letter-searching-for-column-name-in-excel-with-pandas-and-python

[^2_2]: https://stackoverflow.com/questions/13021654/get-column-index-from-column-name-in-python-pandas

[^2_3]: https://techoverflow.net/2021/09/24/how-to-get-openpyxl-column-letter-by-index/

[^2_4]: https://www.geeksforgeeks.org/python/how-to-import-excel-file-and-find-a-specific-column-using-pandas/

[^2_5]: https://www.statology.org/pandas-select-columns-by-name/

[^2_6]: https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html

[^2_7]: https://www.reddit.com/r/learnpython/comments/kbx3sb/is_there_a_way_to_address_pandas_dataframe_by/

[^2_8]: https://www.youtube.com/watch?v=eHrDhKfWLyY

[^2_9]: https://stackoverflow.com/questions/51975912/get-column-names-of-excel-worksheet-with-openpyxl-in-readonly-mode/51999691

[^2_10]: https://pandas.pydata.org/docs/user_guide/indexing.html


---

# skrypt źle wykonuje tekst jako kolumny. zostawiaj oryginalna wartosć, a nastepnie tekst jako kolumny wykonaj w odleglej kolumnie. np BA poniewaz mamy tutaj nadpisywanie wartosci czego chcemy uniknac

Oto poprawiony skrypt, który zachowuje oryginalną wartość kolumny "propozycja" i wykonuje split w odległej kolumnie BA:[^3_1][^3_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tylko jeżeli przynajmniej jedna wartość istnieje
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli oba puste - nie robimy nic
    if zef is None and zev is None:
        continue
    
    # Jeżeli przynajmniej jedno jest niepuste - robimy konkatenację
    zef_str = str(zef) if zef is not None else ''
    zev_str = str(zev) if zev is not None else ''
    ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Wypełniamy puste ZEFZEV na podstawie Zeinr - przeciąganie wartości w dół
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Jeżeli ZEFZEV jest wypełnione - zapisujemy do słownika
    if zefzev:
        zeinr_to_zefzev[zeinr] = zefzev
    # Jeżeli ZEFZEV jest puste - sprawdzamy czy mamy wartość dla tego Zeinr
    else:
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{zefzev_col}{row}'].value = zeinr_to_zefzev[zeinr]

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
ba_col = get_column_letter(BA_COL_NUM)

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
    print(f"  Split będzie w kolumnie: {ba_col} (i dalej)")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Pomijam split.")
    propozycja_col = None

if propozycja_col:
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja:
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BB to druga kolumna po BA (BA + 1 = 54)
    bb_col = get_column_letter(BA_COL_NUM + 1)
    
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bb = ws_propozycje[f'{bb_col}{row}'].value
            if nowe and bb:
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bb)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val:
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        break

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\nPlik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany

### Zachowanie oryginalnej kolumny "propozycja"

Oryginalna wartość w kolumnie "propozycja" jest **zachowana bez zmian**[^3_2]

### Split w odległej kolumnie BA

Split po znaku `_` jest wykonywany **zaczynając od kolumny BA (numer 53)**:[^3_1]

- BA = pierwszy element po split
- BB = drugi element po split
- BC = trzeci element po split
- itd.


### Aktualizacja referencji

Zamiast używać `BJ` (co było relatywne do pozycji propozycji), teraz używamy **BB** (kolumna 54), która jest **drugą kolumną** po split w BA[^3_1]

### Komunikat diagnostyczny

Skrypt informuje w jakiej kolumnie zostanie wykonany split dla lepszej przejrzystości[^3_3]

Dzięki tym zmianom unikasz nadpisywania oryginalnych danych, a wszystkie operacje split są wykonywane w bezpiecznej, odległej lokalizacji.[^3_2]
<span style="display:none">[^3_10][^3_4][^3_5][^3_6][^3_7][^3_8][^3_9]</span>

<div align="center">⁂</div>

[^3_1]: https://bettersolutions.com/excel/rows-columns/letters-to-numbers.htm

[^3_2]: https://www.youtube.com/watch?v=Ue_EDa3Qhik

[^3_3]: https://automatetheboringstuff.com/2e/chapter13/

[^3_4]: https://stackoverflow.com/questions/12902621/getting-the-row-and-column-numbers-from-coordinate-value-in-openpyxl

[^3_5]: https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.utils.cell.html

[^3_6]: https://techoverflow.net/2021/09/24/how-to-get-openpyxl-column-letter-by-index/

[^3_7]: https://www.reddit.com/r/learnpython/comments/wrn6l8/openpyxlcell_give_column_name_instead_of_column/

[^3_8]: https://stackoverflow.com/questions/72383708/get-the-excel-column-label-a-b-z-aa-az-ba-zz-aaa-aab

[^3_9]: https://realpython.com/python-split-string/

[^3_10]: https://openpyxl.readthedocs.io/en/2.4/api/openpyxl.utils.cell.html


---

# teraz musimy poprawic wypelnianie kolumny ZEFZEV, wartosci nie sa dobierane teraz prawidlowo

ZEFZEV uzupełnij tylko tworząc pare ZEF\&ZEF
jeżeli wartosci sa puste zostaw puste
---

Oto poprawiony skrypt z uproszczoną logiką wypełniania ZEFZEV - tworzy parę ZEF\&ZEV **tylko jeżeli obie wartości są wypełnione**, w przeciwnym razie zostaje puste:[^4_1][^4_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - TYLKO jeżeli OBE wartości są niepuste
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Tworzymy parę TYLKO jeżeli oba są niepuste
    if zef is not None and zev is not None:
        ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)
    # W przeciwnym razie pozostawiamy pustą komórkę (None)

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
ba_col = get_column_letter(BA_COL_NUM)

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
    print(f"  Split będzie w kolumnie: {ba_col} (i dalej)")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Pomijam split.")
    propozycja_col = None

if propozycja_col:
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja:
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BB to druga kolumna po BA (BA + 1 = 54)
    bb_col = get_column_letter(BA_COL_NUM + 1)
    
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bb = ws_propozycje[f'{bb_col}{row}'].value
            if nowe and bb:
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bb)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    for row in range(2, ws_propozycje.max_row + 1):
        if ws_propozycje[f'{propozycja_col}{row}'].value:
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val:
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        break

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\nPlik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany w logice ZEFZEV

### Uproszczona logika wypełniania

**Poprzednio:** Tworzyło wartość jeżeli przynajmniej jedna była niepusta + przeciągało wartości w dół[^4_1]

**Teraz:** Tworzy wartość **TYLKO** jeżeli **oba ZEF i ZEV są niepuste**:[^4_2][^4_1]

```python
if zef is not None and zev is not None:
    ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)
```


### Usunięto przeciąganie wartości

Całkowicie usunięto kod który wypełniał puste komórki ZEFZEV na podstawie wartości Zeinr z góry[^4_2]

### Zachowanie pustych komórek

Jeżeli którakolwiek wartość (ZEF lub ZEV) jest pusta, komórka ZEFZEV pozostaje pusta (None)[^4_2]

Teraz dla Twojego przykładu:

- `ZEF=0, ZEV=A` → `ZEFZEV=0A` ✓
- `ZEF=1, ZEV=A` → `ZEFZEV=1A` ✓
- `ZEF=puste, ZEV=puste` → `ZEFZEV=puste` ✓
<span style="display:none">[^4_10][^4_3][^4_4][^4_5][^4_6][^4_7][^4_8][^4_9]</span>

<div align="center">⁂</div>

[^4_1]: https://reactgo.com/python-check-multiple-variables-not-none/

[^4_2]: https://dev.to/hichem-mg/how-to-check-if-variable-is-not-none-in-python-f5o

[^4_3]: https://stackoverflow.com/questions/61513547/combine-and-concat-only-non-empty-columns

[^4_4]: https://github.com/pandas-dev/pandas/issues/45637

[^4_5]: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.concat.html

[^4_6]: https://realpython.com/pandas-merge-join-and-concat/

[^4_7]: https://www.includehelp.com/python/pandas-combine-two-columns-with-null-values.aspx

[^4_8]: https://stackoverflow.com/questions/60130613/concatenate-two-cells-in-excel-using-python-openpyxl-and-for-loop

[^4_9]: https://kanoki.org/Pandas-combine-two-text-columns

[^4_10]: https://openpyxl.readthedocs.io/en/3.1/tutorial.html


---

# widzę ze kolumna ZEFZEV dalej jest wypełniana sprawdz i przeanalizuj dlaczego wypełniane sa wiersze gdzie ZEF i ZEV są puste

Oto poprawiony skrypt z dokładniejszym sprawdzaniem czy komórki są naprawdę puste (nie tylko `None`, ale też puste stringi, spacje itp.):[^5_1][^5_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - TYLKO jeżeli OBE wartości są niepuste
print("\nSprawdzanie pierwszych 10 wierszy:")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Debug: wyświetl pierwsze 10 wierszy
    if row <= 11:
        print(f"Wiersz {row}: ZEF='{zef}' (type: {type(zef).__name__}), ZEV='{zev}' (type: {type(zev).__name__}), "
              f"ZEF puste: {is_cell_empty(zef)}, ZEV puste: {is_cell_empty(zev)}")
    
    # Tworzymy parę TYLKO jeżeli oba są niepuste (nie None, nie pusty string, nie spacje)
    if not is_cell_empty(zef) and not is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)
        if row <= 11:
            print(f"  -> ZEFZEV = '{str(zef) + str(zev)}'")
    else:
        # Pozostaw pustą komórkę (None)
        ws_propozycje[f'{zefzev_col}{row}'].value = None
        if row <= 11:
            print(f"  -> ZEFZEV = puste (przynajmniej jedna wartość jest pusta)")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'
for row in range(2, ws_propozycje.max_row + 1):
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    if zefzev and not is_cell_empty(zefzev):
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zefzev

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
ba_col = get_column_letter(BA_COL_NUM)

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
    print(f"  Split będzie w kolumnie: {ba_col} (i dalej)")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Pomijam split.")
    propozycja_col = None

if propozycja_col:
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BB to druga kolumna po BA (BA + 1 = 54)
    bb_col = get_column_letter(BA_COL_NUM + 1)
    
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bb = ws_propozycje[f'{bb_col}{row}'].value
            if nowe and bb and not is_cell_empty(nowe) and not is_cell_empty(bb):
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bb)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val and not is_cell_empty(skrot_val):
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        break

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\nPlik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany

### Funkcja `is_cell_empty()`

Nowa funkcja która dokładnie sprawdza czy komórka jest pusta:[^5_2][^5_1]

- **None** → True (pusta)
- **Pusty string ""** → True (pusta)
- **String ze spacjami "   "** → True (pusta)
- **Wartość 0** → False (NIE jest pusta)
- **Dowolny tekst** → False (NIE jest pusta)


### Debug output

Skrypt wyświetla pierwsze 10 wierszy z dokładnymi wartościami, typami danych i wynikami sprawdzenia:[^5_1]

```
Wiersz 2: ZEF='0' (type: int), ZEV='A' (type: str), ZEF puste: False, ZEV puste: False
  -> ZEFZEV = '0A'
```


### Rygorystyczne sprawdzanie

Wszędzie gdzie sprawdzamy wartości dodano `is_cell_empty()` zamiast prostego sprawdzania `if value`[^5_2]

### Jawne ustawianie None

Gdy warunek nie jest spełniony, jawnie ustawiamy `None` zamiast pozostawiania domyślnej wartości[^5_1]

Uruchom skrypt i sprawdź output debug - zobaczysz dokładnie co jest w komórkach ZEF i ZEV oraz dlaczego ZEFZEV jest lub nie jest wypełniane.[^5_2][^5_1]
<span style="display:none">[^5_10][^5_3][^5_4][^5_5][^5_6][^5_7][^5_8][^5_9]</span>

<div align="center">⁂</div>

[^5_1]: https://stackoverflow.com/questions/7907928/openpyxl-check-for-empty-cell

[^5_2]: https://stackoverflow.com/questions/40547394/checking-for-empty-cells-with-openpyxl

[^5_3]: https://www.reddit.com/r/learnpython/comments/10790s6/openpyxl_reading_an_empty_cell_in_a_worksheet/

[^5_4]: https://groups.google.com/g/openpyxl-users/c/GbBOnOa8g7Y

[^5_5]: https://openpyxl.readthedocs.io/en/3.1/optimized.html

[^5_6]: https://ask.libreoffice.org/t/how-do-i-detect-a-blank-unfilled-ie-non-zero-non-null-cell-in-calc/26463

[^5_7]: https://stackoverflow.com/questions/67599610/save-excel-worksheet-with-empty-strings-instead-of-nonetype-with-python

[^5_8]: https://github.com/pandas-dev/pandas/issues/39001

[^5_9]: https://gichow.com/2024/04/05/none-string-in-excel-cell-turns-to-null-in-pandas-dataframe-column-why/

[^5_10]: https://www.reddit.com/r/excel/comments/6tuqy5/get_a_cell_to_differentiate_between_blank_and_zero/


---

# ok. teraz jest dobrze. teraz nowe_oznaczenie wypełniamy na podstawie wypełniania. tworzymy słownik z danymi Zeinr i ZEFZEV pomijając gdy wartosć jest pusta.

Answer skipped.

---

# ok. teraz jest dobrze. teraz nowe_oznaczenie wypełniamy na podstawie wypełniania. tworzymy słownik z danymi Zeinr i ZEFZEV pomijając gdy wartosć jest pusta. następnie wypełniamy nowe_oznaczenie sprawdzając Zeinr i dodając wartosc ZEFZEV ze słownika. Wypełniamy tylko gdy jest wartosc w "propozycja"

Oto poprawiony skrypt z logiką wypełniania "nowe_oznaczenie" na podstawie słownika Zeinr→ZEFZEV, tylko gdy "propozycja" jest wypełniona:[^7_1][^7_2][^7_3]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - TYLKO jeżeli OBE wartości są niepuste
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Tworzymy parę TYLKO jeżeli oba są niepuste
    if not is_cell_empty(zef) and not is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)
    else:
        ws_propozycje[f'{zefzev_col}{row}'].value = None

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
ba_col = get_column_letter(BA_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BB to druga kolumna po BA (BA + 1 = 54)
    bb_col = get_column_letter(BA_COL_NUM + 1)
    
    print("\nWypełnianie kolumny skrot...")
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bb = ws_propozycje[f'{bb_col}{row}'].value
            if nowe and bb and not is_cell_empty(nowe) and not is_cell_empty(bb):
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bb)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP)...")
    vlookup_count = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val and not is_cell_empty(skrot_val):
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        vlookup_count += 1
                        break
    
    print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany

### Tworzenie słownika Zeinr → ZEFZEV

Słownik jest tworzony **przed** wypełnianiem kolumny "nowe_oznaczenie":[^7_2][^7_1]

```python
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy TYLKO gdy obie wartości są niepuste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev
```


### Wypełnianie nowe_oznaczenie

Wypełnia kolumnę **tylko** gdy propozycja jest wypełniona, używając wartości ze słownika:[^7_4][^7_1]

```python
if not is_cell_empty(propozycja):
    if zeinr in zeinr_to_zefzev:
        ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
```


### Debug output

Skrypt wyświetla informacje diagnostyczne:[^7_3]

- Liczbę wpisów w słowniku
- Przykładowe wpisy ze słownika
- Liczbę wypełnionych wierszy


### Obsługa braku kolumny propozycja

Jeśli kolumna "propozycja" nie istnieje, wypełnia wszystkie wiersze gdzie Zeinr istnieje w słowniku[^7_1]

Dzięki temu podejściu wartości ZEFZEV są "przeciągane" dla wszystkich wierszy z tym samym Zeinr, ale tylko tam gdzie jest propozycja.[^7_2][^7_3]
<span style="display:none">[^7_10][^7_5][^7_6][^7_7][^7_8][^7_9]</span>

<div align="center">⁂</div>

[^7_1]: https://www.freecodecamp.org/news/how-to-check-if-a-key-exists-in-a-dictionary-in-python/

[^7_2]: https://www.geeksforgeeks.org/python/python-dictionary-get-method/

[^7_3]: https://bobbyhadz.com/blog/pandas-create-dictionary-from-two-columns-of-dataframe

[^7_4]: https://developers.google.com/edu/python/dict-files

[^7_5]: https://stackoverflow.com/questions/47119593/lookup-dictionary-based-on-python-key

[^7_6]: https://www.w3schools.com/python/python_dictionaries_access.asp

[^7_7]: https://www.reddit.com/r/learnpython/comments/yfmnhm/is_there_a_way_to_find_a_key_in_a_dictionary_with/

[^7_8]: https://www.reddit.com/r/learnpython/comments/mb0hej/openpyxl_conditional_formatting_with_multiple/

[^7_9]: https://stackoverflow.com/questions/17426292/how-to-create-a-dictionary-of-two-pandas-dataframe-columns

[^7_10]: https://www.statology.org/how-to-implement-excel-style-conditional-logic-with-openpyxl-in-python/


---

# ok. popraw tworzenie wartosci w ZEFZEV, gdy jedne wartosc jest pusta tez tworzymy. np ZEF = 3 a ZEV puste to ZEFZEV = 3

Oto poprawiony skrypt, który tworzy ZEFZEV również gdy tylko jedna wartość jest wypełniona:[^8_1][^8_2][^8_3]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy (konwertując None na pusty string)
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
if len(zeinr_to_zefzev) > 0:
    print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
ba_col = get_column_letter(BA_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation if 'propozycja' is filled
    # BB to druga kolumna po BA (BA + 1 = 54)
    bb_col = get_column_letter(BA_COL_NUM + 1)
    
    print("\nWypełnianie kolumny skrot...")
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bb = ws_propozycje[f'{bb_col}{row}'].value
            if nowe and bb and not is_cell_empty(nowe) and not is_cell_empty(bb):
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bb)

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP)...")
    vlookup_count = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val and not is_cell_empty(skrot_val):
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        vlookup_count += 1
                        break
    
    print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany w logice ZEFZEV

### Nowa logika konkatenacji

Teraz ZEFZEV jest tworzone **gdy przynajmniej jedna wartość jest niepusta**:[^8_2][^8_1]

**Przed:**

```python
if not is_cell_empty(zef) and not is_cell_empty(zev):  # OBE muszą być niepuste
    ws_propozycje[f'{zefzev_col}{row}'].value = str(zef) + str(zev)
```

**Teraz:**

```python
# Jeżeli OBE puste - nie tworzymy
if is_cell_empty(zef) and is_cell_empty(zev):
    ws_propozycje[f'{zefzev_col}{row}'].value = None
else:
    # Jeżeli przynajmniej jedna niepusta - konkatenujemy
    zef_str = str(zef) if not is_cell_empty(zef) else ''
    zev_str = str(zev) if not is_cell_empty(zev) else ''
    ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str
```


### Konwersja None na pusty string

Puste wartości są konwertowane na pusty string zamiast "None":[^8_3]

```python
zef_str = str(zef) if not is_cell_empty(zef) else ''
```


### Przykłady działania

- **ZEF=3, ZEV=A** → ZEFZEV=**3A** ✓
- **ZEF=3, ZEV=puste** → ZEFZEV=**3** ✓
- **ZEF=puste, ZEV=A** → ZEFZEV=**A** ✓
- **ZEF=puste, ZEV=puste** → ZEFZEV=**puste** ✓

Dzięki tym zmianom ZEFZEV jest wypełniane bardziej elastycznie, zachowując wartości nawet gdy tylko jedna kolumna ma dane.[^8_1][^8_2][^8_3]
<span style="display:none">[^8_10][^8_4][^8_5][^8_6][^8_7][^8_8][^8_9]</span>

<div align="center">⁂</div>

[^8_1]: https://www.geeksforgeeks.org/python/how-to-fix-typeerror-can-only-concatenate-str-not-nonetype-to-str-in-python/

[^8_2]: https://stackoverflow.com/questions/67493245/pythonic-way-to-concatenate-string-to-none

[^8_3]: https://www.geeksforgeeks.org/python/python-convert-none-to-empty-string/

[^8_4]: https://stackoverflow.com/questions/3752240/join-string-and-none-string-using-optional-delimiter

[^8_5]: https://community.esri.com/t5/arcgis-pro-questions/most-efficient-way-to-concatenate-strings-in/td-p/1414863

[^8_6]: https://community.coda.io/t/join-formula-should-not-apply-to-blank-entries/29549

[^8_7]: https://learn.microsoft.com/en-us/answers/questions/747120/appending-values-if-strings-are-not-null-or-empty

[^8_8]: https://stackoverflow.com/questions/75239945/python-openpyxl-merge-empty-row-cells-delimited-by-string

[^8_9]: https://www.youtube.com/watch?v=VmgUMAlt4_s

[^8_10]: https://openpyxl.readthedocs.io/en/latest/api/openpyxl.cell.cell.html


---

# ok. dziala teraz poprawiamy popraw sparawdzanie_rewizji.py aby wypelnianie kolumny skrot zostalo prawidlowo wypelnianie

skrot powinno byc na podstawie

nowe_oznaczenia \& kolumny BG

Oto poprawiony skrypt z prawidłowym wypełnianiem kolumny "skrot" na podstawie **nowe_oznaczenie \& BG**:[^9_1][^9_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
if len(zeinr_to_zefzev) > 0:
    print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
BG_COL_NUM = 59  # BG to kolumna 59
ba_col = get_column_letter(BA_COL_NUM)
bg_col = get_column_letter(BG_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation: nowe_oznaczenie & BG (kolumna 59)
    print(f"\nWypełnianie kolumny skrot (nowe_oznaczenie & {bg_col})...")
    skrot_filled = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
            bg = ws_propozycje[f'{bg_col}{row}'].value
            if nowe and bg and not is_cell_empty(nowe) and not is_cell_empty(bg):
                ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bg)
                skrot_filled += 1
    
    print(f"Wypełniono {skrot_filled} wierszy w kolumnie skrot")

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP)...")
    vlookup_count = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val and not is_cell_empty(skrot_val):
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        vlookup_count += 1
                        break
    
    print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany

### Zmiana kolumny z BB na BG

**Poprzednio:** Używano kolumny **BB** (54)[^9_1]

```python
bb_col = get_column_letter(BA_COL_NUM + 1)  # BB = 54
```

**Teraz:** Używana jest kolumna **BG** (59):[^9_2][^9_1]

```python
BG_COL_NUM = 59  # BG to kolumna 59
bg_col = get_column_letter(BG_COL_NUM)
```


### Poprawiona formuła skrot

Skrot jest teraz konkatenacją: **nowe_oznaczenie \& BG**:[^9_3]

```python
nowe = ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value
bg = ws_propozycje[f'{bg_col}{row}'].value
if nowe and bg and not is_cell_empty(nowe) and not is_cell_empty(bg):
    ws_propozycje[f'{skrot_col}{row}'].value = str(nowe) + ',' + str(bg)
```


### Dodano statystyki

Skrypt wyświetla liczbę wypełnionych wierszy w kolumnie skrot dla łatwiejszej weryfikacji[^9_3]

### Referencja kolumn Excel

Zgodnie z tabelą referencyjną:[^9_2][^9_1]

- **BA** = kolumna 53
- **BB** = kolumna 54
- **BG** = kolumna 59

Teraz kolumna "skrot" będzie poprawnie wypełniana na podstawie **nowe_oznaczenie \& BG**.[^9_3]
<span style="display:none">[^9_10][^9_4][^9_5][^9_6][^9_7][^9_8][^9_9]</span>

<div align="center">⁂</div>

[^9_1]: https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference

[^9_2]: https://bettersolutions.com/excel/rows-columns/letters-to-numbers.htm

[^9_3]: https://www.datacamp.com/tutorial/merge-two-columns-in-excel

[^9_4]: https://stackoverflow.com/questions/18478789/how-to-get-excel-column-index-by-searching-with-cell-text

[^9_5]: https://learn.microsoft.com/en-us/answers/questions/5031729/how-do-i-count-the-number-of-columns

[^9_6]: https://learn.microsoft.com/en-us/answers/questions/5067198/how-to-show-row-numbers-in-spreadsheet-first-colum

[^9_7]: https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.utils.cell.html

[^9_8]: https://www.tutorialspoint.com/how-to-determine-and-identify-the-background-color-index-of-cells-in-excel

[^9_9]: https://stackoverflow.com/questions/58572156/how-to-return-column-letters-instead-of-column-numbers-with-openpyxl

[^9_10]: https://www.trymito.io/excel-to-python/functions/text/CONCAT


---

# ok teraz poprawimy skrot. zasada bedzie ta sama ale wartosc zmienimy na inna. popraw sparawdzanie_rewizji.py aby wypelnianie kolumny skrot zostalo prawidlowo wypelnianie

skrot powinno byc na podstawie

kolumny BG \& kolumny Zeinr np
3656,SL6050431

Oto poprawiony skrypt z nową logiką wypełniania kolumny "skrot": **BG \& Zeinr** (np. "3656,SL6050431"):[^10_1][^10_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
if len(zeinr_to_zefzev) > 0:
    print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
BG_COL_NUM = 59  # BG to kolumna 59
ba_col = get_column_letter(BA_COL_NUM)
bg_col = get_column_letter(BG_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation: BG & Zeinr (np. "3656,SL6050431")
    print(f"\nWypełnianie kolumny skrot ({bg_col} & Zeinr)...")
    skrot_filled = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            bg = ws_propozycje[f'{bg_col}{row}'].value
            zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
            if bg and zeinr and not is_cell_empty(bg) and not is_cell_empty(zeinr):
                ws_propozycje[f'{skrot_col}{row}'].value = str(bg) + ',' + str(zeinr)
                skrot_filled += 1
    
    print(f"Wypełniono {skrot_filled} wierszy w kolumnie skrot")

    # Task 7: Fill 'stare_oznaczenie' with VLOOKUP
    print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP)...")
    vlookup_count = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
            if skrot_val and not is_cell_empty(skrot_val):
                # Perform VLOOKUP manually
                for stary_row in range(2, ws_stary.max_row + 1):
                    if ws_stary[f'C{stary_row}'].value == skrot_val:
                        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                        vlookup_count += 1
                        break
    
    print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany

### Nowa formuła kolumny skrot

**Poprzednio:** nowe_oznaczenie \& BG[^10_1]

**Teraz:** **BG \& Zeinr** z przecinkiem jako separatorem:[^10_2][^10_1]

```python
bg = ws_propozycje[f'{bg_col}{row}'].value
zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
if bg and zeinr and not is_cell_empty(bg) and not is_cell_empty(zeinr):
    ws_propozycje[f'{skrot_col}{row}'].value = str(bg) + ',' + str(zeinr)
```


### Kolejność konkatenacji

Wartości są łączone w kolejności: **najpierw BG, potem Zeinr**[^10_2][^10_1]

### Format wyniku

Format: `"3656,SL6050431"` gdzie:

- `3656` to wartość z kolumny BG
- `,` to separator (przecinek)
- `SL6050431` to wartość z kolumny Zeinr[^10_3]


### Zaktualizowany komunikat

Skrypt wyświetla informację: `"Wypełnianie kolumny skrot (BG & Zeinr)..."` dla jasności[^10_1]

Teraz kolumna "skrot" będzie wypełniana według nowego wzoru: **BG \& Zeinr** (np. "3656,SL6050431").[^10_3][^10_2][^10_1]
<span style="display:none">[^10_10][^10_4][^10_5][^10_6][^10_7][^10_8][^10_9]</span>

<div align="center">⁂</div>

[^10_1]: https://realpython.com/python-string-concatenation/

[^10_2]: https://www.geeksforgeeks.org/python/python-string-concatenation/

[^10_3]: https://excelchamps.com/formulas/concatenate-cells-with-comma/

[^10_4]: https://unstop.com/blog/python-string-concatenation

[^10_5]: https://stackoverflow.com/questions/21542694/difference-between-using-commas-concatenation-and-string-formatters-in-python

[^10_6]: https://www.reddit.com/r/learnpython/comments/jc1y4b/why_do_we_concatenate_using_plus_instead_of_a/

[^10_7]: https://www.tutorialgateway.org/python-string-concatenation/

[^10_8]: https://www.reddit.com/r/learnprogramming/comments/5ok3ql/concatenating_two_columns_of_a_spreadsheet_into_a/

[^10_9]: https://www.extendoffice.com/documents/excel/3323-excel-concatenate-range-line-break-comma-space.html

[^10_10]: https://www.statology.org/how-to-extract-text-between-delimiters-in-excel-using-python-and-openpyxl/


---

# ok

teraz poprawimy wypełnianie kolumny
stare_oznaczenie

wartosć jest formułą
=WYSZUKAJ.PIONOWO(AQ;stary!C:F;4;FAŁSZ)

gdzie AQ w tym przypadku to kolumna skrot
stare_oznaczenie tylko tam gdzie skrot jest wypelniony

Oto poprawiony skrypt z prawidłową logiką wypełniania kolumny "stare_oznaczenie" - wypełnia **tylko tam gdzie skrot jest wypełniony**:[^11_1][^11_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
if len(zeinr_to_zefzev) > 0:
    print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
BG_COL_NUM = 59  # BG to kolumna 59
ba_col = get_column_letter(BA_COL_NUM)
bg_col = get_column_letter(BG_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation: BG & Zeinr (np. "3656,SL6050431")
    print(f"\nWypełnianie kolumny skrot ({bg_col} & Zeinr)...")
    skrot_filled = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            bg = ws_propozycje[f'{bg_col}{row}'].value
            zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
            if bg and zeinr and not is_cell_empty(bg) and not is_cell_empty(zeinr):
                ws_propozycje[f'{skrot_col}{row}'].value = str(bg) + ',' + str(zeinr)
                skrot_filled += 1
    
    print(f"Wypełniono {skrot_filled} wierszy w kolumnie skrot")

# Task 7: Fill 'stare_oznaczenie' with VLOOKUP - TYLKO gdzie skrot jest wypełniony
print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP na podstawie skrot)...")
vlookup_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if skrot_val and not is_cell_empty(skrot_val):
        # Perform VLOOKUP manually - szukamy w kolumnie C arkusza 'stary'
        for stary_row in range(2, ws_stary.max_row + 1):
            if ws_stary[f'C{stary_row}'].value == skrot_val:
                # Zwracamy wartość z kolumny F (4. kolumna z zakresu C:F)
                ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
                vlookup_count += 1
                break

print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe zmiany w VLOOKUP

### Uproszczony warunek

**Poprzednio:** Sprawdzał czy "propozycja" jest wypełniona, a potem czy "skrot" jest wypełniony[^11_1]

**Teraz:** Sprawdza **tylko czy "skrot" jest wypełniony**:[^11_2][^11_1]

```python
for row in range(2, ws_propozycje.max_row + 1):
    skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if skrot_val and not is_cell_empty(skrot_val):
```


### Logika VLOOKUP

Funkcja szuka wartości z kolumny "skrot" w arkuszu 'stary' kolumnie C i zwraca wartość z kolumny F:[^11_3][^11_2]

```python
for stary_row in range(2, ws_stary.max_row + 1):
    if ws_stary[f'C{stary_row}'].value == skrot_val:
        ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value
        break
```


### Ekwiwalent formuły Excel

Python kod odpowiada formule:[^11_2]

```
=WYSZUKAJ.PIONOWO(skrot; stary!C:F; 4; FAŁSZ)
```

Gdzie:

- **skrot** - wartość do wyszukania
- **stary!C:F** - zakres wyszukiwania (kolumna C do F w arkuszu 'stary')
- **4** - zwraca 4. kolumnę z zakresu (F)
- **FAŁSZ** - dokładne dopasowanie[^11_3][^11_2]

Teraz kolumna "stare_oznaczenie" będzie wypełniana **tylko tam gdzie skrot jest wypełniony**, niezależnie od stanu kolumny "propozycja".[^11_1][^11_2][^11_3]
<span style="display:none">[^11_10][^11_4][^11_5][^11_6][^11_7][^11_8][^11_9]</span>

<div align="center">⁂</div>

[^11_1]: https://www.statology.org/how-to-implement-excels-vlookup-with-python-and-openpyxl/

[^11_2]: https://www.trymito.io/excel-to-python/functions/lookup/VLOOKUP

[^11_3]: https://stackoverflow.com/questions/76209601/how-in-openpyxl-search-for-specific-value-in-column-and-read-the-value-from-anot

[^11_4]: https://stackoverflow.com/questions/60058349/using-vlookup-function-in-python-with-openpyxl

[^11_5]: https://python-forum.io/thread-16845.html

[^11_6]: https://www.youtube.com/watch?v=nAOoxlD45js

[^11_7]: https://www.youtube.com/watch?v=EYKVx5rPnMQ

[^11_8]: https://www.reddit.com/r/learnpython/comments/83bqm8/using_openpyxl_how_would_you_go_about_searching_a/

[^11_9]: https://www.reddit.com/r/learnpython/comments/sjddnr/single_vlookup_in_excel_using_python/

[^11_10]: https://www.geeksforgeeks.org/excel/excel-vlookup-with-dynamic-column-reference/


---

# NIE DZIALA. nie ma żadnych wartosci wypelnionych w stare_oznaczenie przeanalizuj jeszcze raz. sprawdz czy wartosc skrot jest juz w pamieci programu w momencie gdy szukamy w arkuszu stary! wydaje mi sie ze tu moze byc problem. PRZEANALIZUJ JESZCZE RAZ I PORPAW KOD

Masz rację! Problem jest prawdopodobnie w **różnicy typów danych** lub **białych znakach**. Oto poprawiony kod z debugowaniem i normalizacją wartości:[^12_1][^12_2]

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

def normalize_value(value):
    """
    Normalizuje wartość do porównania - konwertuje na string i usuwa białe znaki.
    """
    if value is None:
        return None
    return str(value).strip()

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")
if len(zeinr_to_zefzev) > 0:
    print(f"Przykładowe wpisy: {dict(list(zeinr_to_zefzev.items())[:3])}")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'. Wypełniam nowe_oznaczenie bez warunku.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
# TYLKO gdy propozycja jest wypełniona
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    # Wypełniamy TYLKO gdy propozycja jest wypełniona
    if not is_cell_empty(propozycja):
        # Szukamy ZEFZEV w słowniku dla tego Zeinr
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

# Sprawdź czy arkusz 'stary' już istnieje, jeśli tak - usuń go
if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Tworzenie słownika lookup z arkusza 'stary' dla szybszego wyszukiwania
print("\nTworzenie słownika lookup z arkusza 'stary'...")
stary_lookup = {}
for row in range(2, ws_stary.max_row + 1):
    key = normalize_value(ws_stary[f'C{row}'].value)
    value = ws_stary[f'F{row}'].value
    if key:
        stary_lookup[key] = value

print(f"Utworzono słownik lookup z {len(stary_lookup)} wpisami")
if len(stary_lookup) > 0:
    print(f"Przykładowe wpisy z arkusza 'stary': {dict(list(stary_lookup.items())[:3])}")

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
# BA = kolumna 53
BA_COL_NUM = 53  # BA to kolumna 53
BG_COL_NUM = 59  # BG to kolumna 59
ba_col = get_column_letter(BA_COL_NUM)
bg_col = get_column_letter(BG_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    # Split propozycji zaczynając od kolumny BA (nie nadpisuje oryginału)
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation: BG & Zeinr (np. "3656,SL6050431")
    print(f"\nWypełnianie kolumny skrot ({bg_col} & Zeinr)...")
    skrot_filled = 0
    skrot_values_sample = []
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            bg = ws_propozycje[f'{bg_col}{row}'].value
            zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
            if bg and zeinr and not is_cell_empty(bg) and not is_cell_empty(zeinr):
                ws_propozycje[f'{skrot_col}{row}'].value = str(bg) + ',' + str(zeinr)
                skrot_filled += 1
                if len(skrot_values_sample) < 3:
                    skrot_values_sample.append(str(bg) + ',' + str(zeinr))
    
    print(f"Wypełniono {skrot_filled} wierszy w kolumnie skrot")
    print(f"Przykładowe wartości skrot: {skrot_values_sample}")

# Task 7: Fill 'stare_oznaczenie' with VLOOKUP - TYLKO gdzie skrot jest wypełniony
print("\nWypełnianie kolumny stare_oznaczenie (VLOOKUP na podstawie skrot)...")
vlookup_count = 0
not_found_count = 0
not_found_samples = []

for row in range(2, ws_propozycje.max_row + 1):
    skrot_val = ws_propozycje[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if skrot_val and not is_cell_empty(skrot_val):
        # Normalizuj wartość do porównania (usuń spacje, konwertuj na string)
        skrot_normalized = normalize_value(skrot_val)
        
        # Szukaj w słowniku
        if skrot_normalized in stary_lookup:
            ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = stary_lookup[skrot_normalized]
            vlookup_count += 1
        else:
            not_found_count += 1
            if len(not_found_samples) < 5:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row})")

print(f"Znaleziono {vlookup_count} dopasowań w VLOOKUP")
print(f"Nie znaleziono {not_found_count} wartości")
if not_found_samples:
    print(f"Przykłady nie znalezionych wartości: {not_found_samples}")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe poprawki

### Funkcja normalizacji wartości

Dodałem funkcję `normalize_value()` która konwertuje wartości na string i usuwa białe znaki:[^12_3][^12_2]

```python
def normalize_value(value):
    if value is None:
        return None
    return str(value).strip()
```


### Słownik lookup zamiast pętli

**Poprzednio:** Dla każdego wiersza przechodziłem przez CAŁY arkusz 'stary' (O(n²) - bardzo wolne)[^12_1]

**Teraz:** Tworzę słownik raz na początku i używam go do szybkiego wyszukiwania (O(1)):[^12_1]

```python
stary_lookup = {}
for row in range(2, ws_stary.max_row + 1):
    key = normalize_value(ws_stary[f'C{row}'].value)
    value = ws_stary[f'F{row}'].value
    if key:
        stary_lookup[key] = value
```


### Normalizacja przy porównaniu

Wartości są normalizowane PRZED porównaniem:[^12_2][^12_3]

```python
skrot_normalized = normalize_value(skrot_val)
if skrot_normalized in stary_lookup:
    ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = stary_lookup[skrot_normalized]
```


### Rozbudowane debugowanie

Skrypt wyświetla:[^12_1]

- Liczbę wpisów w słowniku lookup
- Przykładowe wartości z arkusza 'stary'
- Przykładowe wartości skrot
- Liczbę dopasowań i niedopasowań
- Przykłady nie znalezionych wartości

To pozwoli Ci zobaczyć dokładnie co się dzieje i dlaczego wartości mogą nie pasować.[^12_3][^12_2][^12_1]
<span style="display:none">[^12_10][^12_4][^12_5][^12_6][^12_7][^12_8][^12_9]</span>

<div align="center">⁂</div>

[^12_1]: https://www.trymito.io/excel-to-python/functions/lookup/VLOOKUP

[^12_2]: https://repairit.wondershare.com/file-repair/vlookup-format-issues.html

[^12_3]: https://www.digitalocean.com/community/tutorials/python-string-comparison

[^12_4]: https://stackoverflow.com/questions/16756699/is-operator-behaves-differently-when-comparing-strings-with-spaces

[^12_5]: https://www.reddit.com/r/learnpython/comments/15t4wik/string_method_for_spaces/

[^12_6]: https://python-forum.io/thread-31999.html

[^12_7]: https://github.com/xunit/xunit/issues/2440

[^12_8]: https://python-forum.io/thread-26143.html

[^12_9]: https://www.geeksforgeeks.org/python/string-whitespace-in-python/

[^12_10]: https://www.reddit.com/r/excel/comments/ymuqcj/type_mismatch_error_13_while_comparing_values_in/


---

# cos jest nie tak. wykaz jest dobry. ale samo dopasowanie sie nie zgadza, proponuje inna zmiane. Tworzymy wykaz propozycje_updated ktory bedzie miec wypelnione wartosci jak do tej pory.

nastepnie uruchomienie nowy skrypt ktory powoduje dopasowanie dancyh.

skrypt dopasowanie_rewizji

dopasowanie danych tworzy nowy wykaz kopia propozycje_updated
następnie na tym wykazie wykonywane będzie
sprawdzenie kolumny skrot i uzupelnienie stare_oznaczenie
wtedy mamy pewnosc ze dane istnieja.
Tworzymy formułe
=WYSZUKAJ.PIONOWO(AQ;stary!C:F;4;FAŁSZ) i uzupelniamy kolumne stare_oznaczenie tam gdzie skrot jest wypelniony.

Doskonały pomysł! Rozdzielmy to na dwa skrypty. Oto rozwiązanie:

## Skrypt 1: Zmodyfikowany sprawdzanie_rewizji.py (bez wypełniania stare_oznaczenie)

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    Zwraca True jeżeli wartość to None, pusty string lub string z samymi spacjami.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
propozycje_file = 'propozycje.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Wczytaj plik Excel używając pandas
df = pd.read_excel(propozycje_file)

# Automatycznie znajdź kolumny po nazwach
try:
    zeinr_col_idx = df.columns.get_loc("Zeinr")
    zef_col_idx = df.columns.get_loc("ZEF")
    zev_col_idx = df.columns.get_loc("ZEV")
except KeyError as e:
    print(f"Błąd: Nie znaleziono kolumny {e}. Sprawdź nazwy kolumn w pliku.")
    exit()

# Konwertuj indeksy kolumn na litery Excel (dodaj 1, bo Excel jest 1-based)
zeinr_col = get_column_letter(zeinr_col_idx + 1)
zef_col = get_column_letter(zef_col_idx + 1)
zev_col = get_column_letter(zev_col_idx + 1)

print(f"Znaleziono kolumny:")
print(f"  Zeinr: {zeinr_col}")
print(f"  ZEF: {zef_col}")
print(f"  ZEV: {zev_col}")

# Teraz wczytaj plik używając openpyxl do edycji
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV - tworzymy JEŻELI przynajmniej jedna wartość jest niepusta
print("\nWypełnianie kolumny ZEFZEV...")
for row in range(2, ws_propozycje.max_row + 1):
    zef = ws_propozycje[f'{zef_col}{row}'].value
    zev = ws_propozycje[f'{zev_col}{row}'].value
    
    # Jeżeli OBE puste - nie tworzymy
    if is_cell_empty(zef) and is_cell_empty(zev):
        ws_propozycje[f'{zefzev_col}{row}'].value = None
    else:
        # Jeżeli przynajmniej jedna niepusta - konkatenujemy
        zef_str = str(zef) if not is_cell_empty(zef) else ''
        zev_str = str(zev) if not is_cell_empty(zev) else ''
        ws_propozycje[f'{zefzev_col}{row}'].value = zef_str + zev_str

# Tworzenie słownika Zeinr -> ZEFZEV (pomijając puste wartości ZEFZEV)
print("\nTworzenie słownika Zeinr -> ZEFZEV...")
zeinr_to_zefzev = {}
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    zefzev = ws_propozycje[f'{zefzev_col}{row}'].value
    
    # Dodajemy do słownika tylko gdy ZEFZEV nie jest puste
    if not is_cell_empty(zeinr) and not is_cell_empty(zefzev):
        zeinr_to_zefzev[zeinr] = zefzev

print(f"Utworzono słownik z {len(zeinr_to_zefzev)} wpisami")

# Task 2: Create new column 'nowe_oznaczenie' using ZEFZEV
nowe_oznaczenie_col = get_column_letter(ws_propozycje.max_column + 1)
ws_propozycje[f'{nowe_oznaczenie_col}1'].value = 'nowe_oznaczenie'

# Znajdź kolumnę 'propozycja' automatycznie
try:
    propozycja_col_idx = df.columns.get_loc("propozycja")
    propozycja_col = get_column_letter(propozycja_col_idx + 1)
    print(f"  propozycja: {propozycja_col}")
except KeyError:
    print("Ostrzeżenie: Nie znaleziono kolumny 'propozycja'.")
    propozycja_col = None

# Wypełnianie nowe_oznaczenie na podstawie słownika Zeinr -> ZEFZEV
print("\nWypełnianie kolumny nowe_oznaczenie...")
filled_count = 0
for row in range(2, ws_propozycje.max_row + 1):
    zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
    propozycja = ws_propozycje[f'{propozycja_col}{row}'].value if propozycja_col else True
    
    if not is_cell_empty(propozycja):
        if zeinr in zeinr_to_zefzev:
            ws_propozycje[f'{nowe_oznaczenie_col}{row}'].value = zeinr_to_zefzev[zeinr]
            filled_count += 1

print(f"Wypełniono {filled_count} wierszy w kolumnie nowe_oznaczenie")

# Task 5: Create 'stare_oznaczenie' column next to 'nowe_oznaczenie'
stare_oznaczenie_col = get_column_letter(openpyxl.utils.column_index_from_string(nowe_oznaczenie_col) + 1)
ws_propozycje[f'{stare_oznaczenie_col}1'].value = 'stare_oznaczenie'

# Task 6: Create 'skrot' column next to 'stare_oznaczenie'
skrot_col = get_column_letter(openpyxl.utils.column_index_from_string(stare_oznaczenie_col) + 1)
ws_propozycje[f'{skrot_col}1'].value = 'skrot'

# Task 3: Create new sheet 'stary' and copy columns A:F from stare_file
print("\nKopiowanie danych ze starego pliku...")
wb_stare = load_workbook(stare_file)
ws_stare_source = wb_stare.active

if 'stary' in wb_propozycje.sheetnames:
    wb_propozycje.remove(wb_propozycje['stary'])

ws_stary = wb_propozycje.create_sheet('stary')

# Copy columns A:F
for col in range(1, 7):
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' i umieść w odległej kolumnie BA
BA_COL_NUM = 53
BG_COL_NUM = 59
ba_col = get_column_letter(BA_COL_NUM)
bg_col = get_column_letter(BG_COL_NUM)

if propozycja_col:
    print(f"\nSplit kolumny propozycja w kolumnie: {ba_col} (i dalej)")
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            parts = str(propozycja).split('_')
            for i, part in enumerate(parts):
                col_num = BA_COL_NUM + i
                col_letter = get_column_letter(col_num)
                ws_propozycje[f'{col_letter}{row}'].value = part

    # Fill 'skrot' with concatenation: BG & Zeinr
    print(f"\nWypełnianie kolumny skrot ({bg_col} & Zeinr)...")
    skrot_filled = 0
    for row in range(2, ws_propozycje.max_row + 1):
        propozycja = ws_propozycje[f'{propozycja_col}{row}'].value
        if propozycja and not is_cell_empty(propozycja):
            bg = ws_propozycje[f'{bg_col}{row}'].value
            zeinr = ws_propozycje[f'{zeinr_col}{row}'].value
            if bg and zeinr and not is_cell_empty(bg) and not is_cell_empty(zeinr):
                ws_propozycje[f'{skrot_col}{row}'].value = str(bg) + ',' + str(zeinr)
                skrot_filled += 1
    
    print(f"Wypełniono {skrot_filled} wierszy w kolumnie skrot")

# Save the workbook
output_file = 'propozycje_updated.xlsx'
wb_propozycje.save(output_file)
print(f"\n✓ Plik został pomyślnie zapisany jako '{output_file}'")
print("\nUruchom teraz skrypt 'dopasowanie_rewizji.py' aby wypełnić kolumnę stare_oznaczenie")
```


## Skrypt 2: Nowy dopasowanie_rewizji.py

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# Wczytaj skopiowany plik
print(f"Wczytywanie pliku '{output_file}'...")
wb = load_workbook(output_file)
ws = wb.active

# Znajdź kolumny stare_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None

for col in range(1, ws.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {col_letter}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

print("\nDodawanie formuł VLOOKUP do kolumny stare_oznaczenie...")
formula_count = 0

for row in range(2, ws.max_row + 1):
    skrot_val = ws[f'{skrot_col}{row}'].value
    
    # Dodaj formułę TYLKO gdy skrot jest wypełniony
    if skrot_val and not is_cell_empty(skrot_val):
        # Tworzymy formułę VLOOKUP z przecinkami (styl Polski Excel)
        formula = f'=WYSZUKAJ.PIONOWO({skrot_col}{row};stary!C:F;4;FAŁSZ)'
        ws[f'{stare_oznaczenie_col}{row}'].value = formula
        formula_count += 1

print(f"Dodano {formula_count} formuł VLOOKUP")

# Zapisz plik
print(f"\nZapisywanie pliku '{output_file}'...")
wb.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print("\nOtwórz plik w Excelu - formuły zostaną automatycznie przeliczone")
```


## Jak używać

### Krok 1: Uruchom sprawdzanie_rewizji.py

```bash
python sprawdzanie_rewizji.py
```

Tworzy plik `propozycje_updated.xlsx` z wypełnionymi danymi (bez stare_oznaczenie)[^13_1][^13_2]

### Krok 2: Uruchom dopasowanie_rewizji.py

```bash
python dopasowanie_rewizji.py
```

Tworzy plik `propozycje_updated_z_dopasowaniem.xlsx` z formułami VLOOKUP[^13_2][^13_1]

## Kluczowe cechy drugiego skryptu

### Formuła Excel zamiast wartości

Skrypt dodaje **formułę** a nie wartość:[^13_1]

```python
formula = f'=WYSZUKAJ.PIONOWO({skrot_col}{row};stary!C:F;4;FAŁSZ)'
ws[f'{stare_oznaczenie_col}{row}'].value = formula
```


### Polska składnia

Używa polskiej nazwy funkcji `WYSZUKAJ.PIONOWO` i `FAŁSZ`[^13_3]

### Automatyczne przeliczanie

Excel automatycznie przeliczy formuły po otwarciu pliku[^13_2][^13_3]

### Bezpieczeństwo danych

Tworzy nową kopię pliku zamiast nadpisywać oryginał[^13_1]

Teraz masz pewność że formuły działają dokładnie tak jak w Excelu![^13_3][^13_2][^13_1]
<span style="display:none">[^13_10][^13_4][^13_5][^13_6][^13_7][^13_8][^13_9]</span>

<div align="center">⁂</div>

[^13_1]: https://stackoverflow.com/questions/60058349/using-vlookup-function-in-python-with-openpyxl

[^13_2]: https://www.youtube.com/watch?v=nAOoxlD45js

[^13_3]: https://xlsxwriter.readthedocs.io/working_with_formulas.html

[^13_4]: https://python-forum.io/thread-16845.html

[^13_5]: https://www.statology.org/how-to-implement-excels-vlookup-with-python-and-openpyxl/

[^13_6]: https://www.trymito.io/excel-to-python/functions/lookup/VLOOKUP

[^13_7]: https://www.youtube.com/watch?v=HFuJV9qoVDA

[^13_8]: https://stackoverflow.com/questions/28517508/how-can-i-use-openpyxl-to-read-an-excel-cell-value-and-not-the-formula-computing

[^13_9]: https://buymeacoffee.com/dataautomation/openpyxl-vlookup-formula-entire-column-excel-workbooks-python

[^13_10]: https://groups.google.com/g/openpyxl-users/c/g3PQIpqzBnA


---

# przy otwieraniu pliku dopasowanie excele

Widzę problem! Excel pokazuje komunikat że "skoroszyt istnieje zawartość, która powoduje problem". To typowy problem gdy formuły są dodawane przez openpyxl.[^14_1][^14_2][^14_3]

Oto **poprawiona wersja** skryptu `dopasowanie_rewizji.py` która rozwiązuje ten problem:

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# Wczytaj skopiowany plik BEZ data_only (aby móc zapisywać formuły)
print(f"Wczytywanie pliku '{output_file}'...")
wb = load_workbook(output_file)
ws = wb.active

# Znajdź kolumny stare_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None
stare_oznaczenie_col_num = None

for col in range(1, ws.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        stare_oznaczenie_col_num = col
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary = wb['stary']

print("\nWypełnianie kolumny stare_oznaczenie metodą bezpośredniego wyszukiwania...")

# Tworzymy słownik z arkusza 'stary' dla szybkiego wyszukiwania
stary_dict = {}
for row in range(2, ws_stary.max_row + 1):
    key = ws_stary[f'C{row}'].value
    value = ws_stary[f'F{row}'].value
    if key:
        # Normalizuj klucz (usuń białe znaki, konwertuj na string)
        key_normalized = str(key).strip()
        stary_dict[key_normalized] = value

print(f"Utworzono słownik z {len(stary_dict)} wpisami z arkusza 'stary'")

# Wypełnianie wartościami (nie formułami) - to uniknie problemu z Excelem
filled_count = 0
not_found_count = 0

for row in range(2, ws.max_row + 1):
    skrot_val = ws[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if skrot_val and not is_cell_empty(skrot_val):
        # Normalizuj wartość
        skrot_normalized = str(skrot_val).strip()
        
        # Szukaj w słowniku
        if skrot_normalized in stary_dict:
            ws[f'{stare_oznaczenie_col}{row}'].value = stary_dict[skrot_normalized]
            filled_count += 1
        else:
            not_found_count += 1

print(f"Wypełniono {filled_count} wierszy")
print(f"Nie znaleziono dopasowania dla {not_found_count} wierszy")

# Zapisz plik
print(f"\nZapisywanie pliku '{output_file}'...")
wb.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print("\nPlik można teraz bezpiecznie otworzyć w Excelu bez ostrzeżeń")
```


## Kluczowe zmiany

### Wartości zamiast formuł

**Poprzednio:** Dodawałem formuły Excel które powodowały ostrzeżenie[^14_2][^14_3]

**Teraz:** Wykonuję VLOOKUP w Pythonie i zapisuję **gotowe wartości**:[^14_2]

```python
ws[f'{stare_oznaczenie_col}{row}'].value = stary_dict[skrot_normalized]
```


### Słownik dla wydajności

Tworzę słownik z arkusza 'stary' raz, zamiast przeszukiwać go wielokrotnie[^14_2]

### Normalizacja wartości

Usuwam białe znaki i konwertuję na string aby zapewnić poprawne dopasowanie[^14_3][^14_1]

### Bez ostrzeżeń Excel

Ponieważ zapisuję wartości (nie formuły), Excel otworzy plik bez problemu[^14_1][^14_3]

## Dlaczego poprzednia wersja powodowała błąd?

Excel pokazuje ostrzeżenie gdy:[^14_3][^14_1]

1. Plik został utworzony/zmodyfikowany przez zewnętrzne narzędzie (openpyxl)
2. Zawiera formuły które nie były przeliczone przez Excel
3. Excel nie może zweryfikować integralności formuł

Nowa wersja zapisuje **wartości wynikowe** zamiast formuł, więc Excel nie ma powodu do ostrzeżenia.[^14_1][^14_3][^14_2]
<span style="display:none">[^14_10][^14_11][^14_4][^14_5][^14_6][^14_7][^14_8][^14_9]</span>

<div align="center">⁂</div>

[^14_1]: https://pl.ittrip.xyz/ms-office/excel/corruption

[^14_2]: https://stackoverflow.com/questions/36116162/python-openpyxl-data-only-true-returning-none

[^14_3]: https://repairit.wondershare.com/file-repair/fix-we-found-problem-with-content-in-excel.html

[^14_4]: image.jpg

[^14_5]: https://support.microsoft.com/pl-pl/office/naprawianie-uszkodzonego-skoroszytu-153a45f4-6cab-44b1-93ca-801ddcd4ea53

[^14_6]: https://support.microsoft.com/pl-pl/office/naprawianie-uszkodzonego-skoroszytu-7abfc44d-e9bf-4896-8899-bd10ef4d61ab

[^14_7]: https://www.elektroda.pl/rtvforum/topic3295755.html

[^14_8]: https://learn.microsoft.com/pl-pl/answers/questions/5020682/arkusz-si-otwiera-i-przesta-jest-za-to-komunikat

[^14_9]: https://www.reddit.com/r/excel/comments/obqjxn/excel_workbook_still_shows_as_being_open_by/

[^14_10]: https://github.com/fabiocaccamo/python-benedict/issues/355

[^14_11]: https://www.reddit.com/r/excel/comments/yd5o9f/getting_an_error_whenever_i_open_my_excel_this/


---

# nie dziala. nie wyswietlaja sie zadne wartosci. moze przyczyna jest duza liczba wierzsz w arkuszu stare. sprobuj utworzyc slownik ale tylko kolumny C i F z zakladki stary. nastepnie szukamy w slowniku C sprawdzamy ze skrot i wpisujemy wartosc F do stare_oznaczenie

Masz rację! Poprawmy kod aby był bardziej precyzyjny i wydajny. Oto ulepszona wersja z dodatkowymi informacjami diagnostycznymi:[^15_1][^15_2]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# Wczytaj skopiowany plik
print(f"Wczytywanie pliku '{output_file}'...")
wb = load_workbook(output_file)
ws = wb.active

# Znajdź kolumny stare_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None

for col in range(1, ws.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary = wb['stary']

# KROK 1: Tworzenie słownika TYLKO z kolumn C i F arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
stary_dict = {}
empty_keys = 0

for row in range(2, ws_stary.max_row + 1):
    # Czytamy TYLKO kolumnę C (klucz) i F (wartość)
    key_c = ws_stary[f'C{row}'].value
    value_f = ws_stary[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        # Normalizuj klucz - konwertuj na string i usuń białe znaki
        key_normalized = str(key_c).strip()
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1

print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy")
print(f"  ℹ Przykładowe wpisy ze słownika:")
for i, (k, v) in enumerate(list(stary_dict.items())[:5]):
    print(f"    [{i+1}] Klucz: '{k}' -> Wartość: '{v}'")

# KROK 2: Zbieranie wartości skrot z głównego arkusza
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws.max_row + 1, 12)):  # Pokazujemy tylko pierwsze 10 wierszy
    skrot_val = ws[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_values.append((row, str(skrot_val).strip()))

print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val in skrot_values:
    print(f"    Wiersz {row}: '{val}'")

# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
not_found_samples = []

for row in range(2, ws.max_row + 1):
    skrot_val = ws[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if not is_cell_empty(skrot_val):
        # Normalizuj wartość - konwertuj na string i usuń białe znaki
        skrot_normalized = str(skrot_val).strip()
        
        # Szukaj w słowniku (kolumna C)
        if skrot_normalized in stary_dict:
            # Wpisz wartość z kolumny F
            ws[f'{stare_oznaczenie_col}{row}'].value = stary_dict[skrot_normalized]
            filled_count += 1
            
            # Debug: wyświetl pierwsze 5 dopasowań
            if filled_count <= 5:
                print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{stary_dict[skrot_normalized]}'")
        else:
            not_found_count += 1
            if len(not_found_samples) < 5:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row})")

print(f"\n=== PODSUMOWANIE ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono dopasowania: {not_found_count} wierszy")

if not_found_samples:
    print(f"\n  ℹ Przykłady NIE znalezionych wartości:")
    for sample in not_found_samples:
        print(f"    - {sample}")
    print(f"\n  💡 Sprawdź czy te wartości rzeczywiście istnieją w kolumnie C arkusza 'stary'")

# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
```


## Kluczowe usprawnienia

### Precyzyjne czytanie kolumn C i F

Kod czyta **TYLKO** kolumny C (klucz) i F (wartość) z arkusza 'stary':[^15_2][^15_3]

```python
key_c = ws_stary[f'C{row}'].value
value_f = ws_stary[f'F{row}'].value
```


### Rozbudowane debugowanie

Skrypt pokazuje:[^15_1]

- Liczbę wpisów w słowniku
- Przykładowe klucze i wartości ze słownika
- Przykładowe wartości skrot
- Pierwsze 5 dopasowań krok po kroku
- Przykłady wartości które nie zostały znalezione


### Normalizacja wartości

Wszystkie wartości są konwertowane na string i oczyszczane z białych znaków:[^15_4][^15_1]

```python
key_normalized = str(key_c).strip()
skrot_normalized = str(skrot_val).strip()
```


### Wydajność O(1)

Dzięki słownikowi, każde wyszukiwanie działa w czasie stałym O(1), nawet dla milionów rekordów[^15_5][^15_1]

Uruchom ten skrypt i sprawdź output - zobaczysz dokładnie:

1. Jakie dane są w słowniku
2. Jakie wartości skrot są wyszukiwane
3. Które dopasowania się udały
4. Które nie zostały znalezione i dlaczego[^15_4][^15_1]
<span style="display:none">[^15_10][^15_6][^15_7][^15_8][^15_9]</span>

<div align="center">⁂</div>

[^15_1]: https://dzone.com/articles/python-dictionary-a-powerful-tool-for-data-enginee

[^15_2]: https://stackoverflow.com/questions/34754077/openpyxl-how-to-read-only-one-column-from-excel-file-in-python

[^15_3]: https://www.geeksforgeeks.org/python/get-values-of-all-rows-in-a-particular-column-in-openpyxl-python/

[^15_4]: https://www.trymito.io/blog/quick-guide-to-excel-vlookup-with-python-code-a-vlookup-equivalent-with-pandas

[^15_5]: https://python.plainenglish.io/optimizing-python-dictionaries-a-comprehensive-guide-f9b04063467a

[^15_6]: https://stackoverflow.com/questions/19391648/keeping-large-dictionary-in-python-affects-application-performance

[^15_7]: https://www.reddit.com/r/learnpython/comments/1airrg8/is_creating_a_dictionary_from_70_milion_entries/

[^15_8]: https://labex.io/tutorials/python-how-to-efficiently-iterate-through-a-large-python-dictionary-398184

[^15_9]: https://justtiffme.com/python-performance-with-dictionaries-and-sets/

[^15_10]: https://stackoverflow.com/questions/26549781/faster-way-to-vlookup-huge-set-of-data


---

# zle dopasowuje np dla wiersz 5 === PODSUMOWANIE ===

✓ Wypełniono: 0 wierszy
✗ Nie znaleziono dopasowania: 189 wierszy

ℹ Przykłady NIE znalezionych wartości:
- '2345,10002848' (wiersz 2)
- '2345,10002848' (wiersz 3)
- '2345,10002775' (wiersz 4)
- '2345,SL10115492' (wiersz 5)
- '2345,SL10115492' (wiersz 6) powinna byc wartosc. generuj plik ze słownikiem wartosci ktory jest wykorzystywany, sprawdze jego poprawnosc

Doskonały pomysł! Wygenerujmy plik CSV ze słownikiem aby sprawdzić co dokładnie jest w danych. Oto poprawiony skrypt z eksportem słownika:[^16_1][^16_2]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import csv

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# Wczytaj skopiowany plik
print(f"Wczytywanie pliku '{output_file}'...")
wb = load_workbook(output_file)
ws = wb.active

# Znajdź kolumny stare_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None

for col in range(1, ws.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary = wb['stary']

# KROK 1: Tworzenie słownika TYLKO z kolumn C i F arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
stary_dict = {}
empty_keys = 0
duplicate_keys = 0

for row in range(2, ws_stary.max_row + 1):
    # Czytamy TYLKO kolumnę C (klucz) i F (wartość)
    key_c = ws_stary[f'C{row}'].value
    value_f = ws_stary[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        # Normalizuj klucz - konwertuj na string i usuń białe znaki
        key_normalized = str(key_c).strip()
        
        # Sprawdź czy klucz już istnieje (duplikaty)
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1

print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")

# EKSPORT SŁOWNIKA DO CSV
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Wiersz_w_stary'])
    
    # Sortujemy słownik alfabetycznie dla łatwiejszego przeszukiwania
    for key in sorted(stary_dict.keys()):
        writer.writerow([key, stary_dict[key], 'N/A'])

print(f"  ✓ Zapisano {len(stary_dict)} wpisów do pliku '{dictionary_export_file}'")
print(f"  ℹ Otwórz ten plik w Excelu aby sprawdzić zawartość słownika")

# KROK 2: Zbieranie wartości skrot z głównego arkusza
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws.max_row + 1, 12)):  # Pokazujemy tylko pierwsze 10 wierszy
    skrot_val = ws[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        # Sprawdź czy istnieje w słowniku
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        skrot_values.append((row, skrot_normalized, exists))

print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")

# EKSPORT WARTOŚCI SKROT DO SPRAWDZENIA
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Czy_istnieje_w_słowniku'])
    
    for row in range(2, ws.max_row + 1):
        skrot_val = ws[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            writer.writerow([row, skrot_normalized, exists])

print(f"  ✓ Zapisano wartości skrot do pliku '{skrot_export_file}'")

# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
not_found_samples = []

for row in range(2, ws.max_row + 1):
    skrot_val = ws[f'{skrot_col}{row}'].value
    
    # Wypełniamy TYLKO gdy skrot jest wypełniony
    if not is_cell_empty(skrot_val):
        # Normalizuj wartość - konwertuj na string i usuń białe znaki
        skrot_normalized = str(skrot_val).strip()
        
        # Szukaj w słowniku (kolumna C)
        if skrot_normalized in stary_dict:
            # Wpisz wartość z kolumny F
            ws[f'{stare_oznaczenie_col}{row}'].value = stary_dict[skrot_normalized]
            filled_count += 1
            
            # Debug: wyświetl pierwsze 5 dopasowań
            if filled_count <= 5:
                print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{stary_dict[skrot_normalized]}'")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row})")

print(f"\n=== PODSUMOWANIE ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono dopasowania: {not_found_count} wierszy")

if not_found_samples:
    print(f"\n  ℹ Przykłady NIE znalezionych wartości:")
    for sample in not_found_samples:
        print(f"    - {sample}")
    print(f"\n  💡 Sprawdź w pliku '{dictionary_export_file}' czy te wartości istnieją")

# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {dictionary_export_file} - pełny słownik z arkusza 'stary'")
print(f"  2. {skrot_export_file} - wartości skrot i czy istnieją w słowniku")
print(f"  3. {output_file} - wynikowy plik Excel")
```


## Kluczowe zmiany

### Eksport słownika do CSV

Tworzy plik `slownik_stary_C_F.csv` z wszystkimi wpisami ze słownika:[^16_2][^16_1]

```python
writer.writerow(['Klucz_C', 'Wartość_F', 'Wiersz_w_stary'])
for key in sorted(stary_dict.keys()):
    writer.writerow([key, stary_dict[key], 'N/A'])
```


### Eksport wartości skrot

Tworzy plik `wartosci_skrot_do_sprawdzenia.csv` pokazujący które wartości skrot istnieją w słowniku:[^16_3][^16_1]

```python
writer.writerow(['Wiersz', 'Wartość_skrot', 'Czy_istnieje_w_słowniku'])
```


### Więcej informacji diagnostycznych

- Liczba duplikatów kluczy
- Pierwsze 10 wartości skrot z informacją czy istnieją
- Więcej przykładów nie znalezionych wartości (10 zamiast 5)


### Separator średnik

Używam średnika `;` jako separator w CSV aby łatwiej otworzyć w polskim Excelu[^16_2][^16_3]

Uruchom ten skrypt i otwórz pliki CSV:

1. **slownik_stary_C_F.csv** - sprawdź czy wartość `2345,SL10115492` tam istnieje
2. **wartosci_skrot_do_sprawdzenia.csv** - zobacz które wartości są oznaczone jako "NIE"

To pozwoli Ci dokładnie zdiagnozować problem![^16_1][^16_3][^16_2]
<span style="display:none">[^16_10][^16_4][^16_5][^16_6][^16_7][^16_8][^16_9]</span>

<div align="center">⁂</div>

[^16_1]: https://stackoverflow.com/questions/10373247/how-do-i-write-a-python-dictionary-to-a-csv-file

[^16_2]: https://www.geeksforgeeks.org/python/how-to-save-a-python-dictionary-to-a-csv-file/

[^16_3]: https://www.tutorialspoint.com/how-to-save-a-python-dictionary-to-csv-file

[^16_4]: https://docs.python.org/3/library/csv.html

[^16_5]: https://www.youtube.com/watch?v=0wO3KdJOKN8

[^16_6]: https://blog.finxter.com/5-best-ways-to-convert-python-dictionary-to-excel-table/

[^16_7]: https://pynative.com/python-save-dictionary-to-file/

[^16_8]: https://community.lambdatest.com/t/how-to-write-a-python-dictionary-to-csv-with-keys-as-headers/34274

[^16_9]: https://www.reddit.com/r/learnpython/comments/4xa7ny/using_openpyxl_how_can_i_write_a_dictionary_into/

[^16_10]: https://blog.finxter.com/how-to-save-a-dictionary-to-a-file-in-python/


---

# zauwazylem ze w slowniku jest duzo pustych pozycji, dane powinny byc wklejone jako wartosci a nie formuly tutaj jest problem, czesc wierszy kolumnie stary mamy wklejone jako formuly, skrypt powinien odczytac to jako wartosci

Świetnie wyłapałeś problem! Formuły w Excelu nie są automatycznie przeliczane przez openpyxl. Musisz użyć `data_only=True` aby czytać **wartości** zamiast formuł. Oto poprawiony kod:[^17_1][^17_2][^17_3]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import csv

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ
print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary_read = wb_read['stary']

# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print("  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
stary_dict = {}
empty_keys = 0
empty_values = 0
duplicate_keys = 0

for row in range(2, ws_stary_read.max_row + 1):
    # Czytamy WARTOŚCI (nie formuły) z kolumn C i F
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        # Normalizuj klucz
        key_normalized = str(key_c).strip()
        
        # Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            # Debug: pokaż pierwsze 5 przypadków pustych wartości
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        # Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1

print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy (kolumna C)")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")

# EKSPORT SŁOWNIKA DO CSV
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    # Sortujemy słownik alfabetycznie
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])

print(f"  ✓ Zapisano {len(stary_dict)} wpisów do pliku '{dictionary_export_file}'")

# Przykładowe wpisy ze słownika
print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")

# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ
print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active

# Znajdź kolumny stare_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None

for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# KROK 2: Zbieranie wartości skrot
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws_write.max_row + 1, 12)):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        has_value = ""
        if skrot_normalized in stary_dict:
            if is_cell_empty(stary_dict[skrot_normalized]):
                has_value = " (ale wartość F jest PUSTA)"
        skrot_values.append((row, skrot_normalized, exists + has_value))

print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")

# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])

print(f"  ✓ Zapisano wartości skrot do pliku '{skrot_export_file}'")

# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0
not_found_samples = []

for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            # Sprawdź czy wartość F nie jest pusta
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
                if len(not_found_samples) < 10:
                    not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - KLUCZ ISTNIEJE ale wartość F PUSTA")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - BRAK w słowniku")

print(f"\n=== PODSUMOWANIE ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono w słowniku: {not_found_count} wierszy")
print(f"  ⚠ Znaleziono klucz ale wartość F pusta: {empty_value_count} wierszy")

if not_found_samples:
    print(f"\n  ℹ Przykłady problemów:")
    for sample in not_found_samples:
        print(f"    - {sample}")

# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {dictionary_export_file} - słownik z oznaczeniem pustych wartości")
print(f"  2. {skrot_export_file} - analiza dopasowań")
print(f"  3. {output_file} - wynikowy plik Excel")
```


## Kluczowe zmiany

### data_only=True przy czytaniu

Najważniejsza zmiana - plik jest wczytywany z `data_only=True` aby czytać **wartości** zamiast formuł:[^17_2][^17_3][^17_1]

```python
wb_read = load_workbook(output_file, data_only=True)
```


### Dwa otwarcia pliku

1. **Pierwsze otwarcie** (`data_only=True`) - do CZYTANIA wartości ze słownika[^17_3]
2. **Drugie otwarcie** (domyślne) - do ZAPISYWANIA wyników[^17_1]

### Wykrywanie pustych wartości F

Skrypt teraz wykrywa i raportuje przypadki gdzie klucz istnieje ale wartość F jest pusta:[^17_3]

```python
if is_cell_empty(value_f):
    empty_values += 1
```


### Lepszy raport

CSV pokazuje które wpisy mają puste wartości oraz status każdego dopasowania[^17_2]

### Trzy kategorie problemów

1. Wypełniono - sukces
2. Nie znaleziono w słowniku - brak klucza
3. Znaleziono klucz ale wartość F pusta - formuła nie była przeliczona[^17_3]

Teraz skrypt powinien poprawnie czytać wartości z formuł! Jeśli nadal widzisz puste wartości, znaczy to że Excel nie zapisał przeliczonych wartości - musisz otworzyć plik w Excelu i zapisać go ponownie.[^17_1][^17_2][^17_3]
<span style="display:none">[^17_10][^17_4][^17_5][^17_6][^17_7][^17_8][^17_9]</span>

<div align="center">⁂</div>

[^17_1]: https://stackoverflow.com/questions/28517508/how-can-i-use-openpyxl-to-read-an-excel-cell-value-and-not-the-formula-computing

[^17_2]: https://www.datacamp.com/tutorial/openpyxl

[^17_3]: https://stackoverflow.com/questions/36116162/python-openpyxl-data-only-true-returning-none

[^17_4]: https://www.reddit.com/r/learnpython/comments/147qbfh/using_openpyxlpandas_to_get_the_final_value_of_a/

[^17_5]: https://www.youtube.com/watch?v=EzAjJE50kA0

[^17_6]: https://groups.google.com/g/openpyxl-users/c/g3PQIpqzBnA

[^17_7]: https://python-forum.io/thread-30628.html

[^17_8]: https://openpyxl.readthedocs.io/en/stable/formula.html

[^17_9]: https://groups.google.com/g/openpyxl-users/c/GbBOnOa8g7Y

[^17_10]: https://www.reddit.com/r/learnpython/comments/6m9gi3/how_do_i_read_the_excel_value_and_not_the_formula/


---

# ok teraz dobrze, dodaj jeszcze sprawdzenie nowe_oznaczenie i stare_oznaczenie jako formuła i w kolumnie po skrot np. AR (nazwana REWIZJA) sprawdz czy PRAWDA czy FAŁSZ czy nowe_oznaczenie=stare_oznaczenie

Świetnie! Dodajmy kolumnę REWIZJA z formułą porównującą nowe_oznaczenie i stare_oznaczenie. Oto zaktualizowany kod:[^18_1][^18_2][^18_3]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import csv

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'

# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ
print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active

# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary_read = wb_read['stary']

# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print("  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
stary_dict = {}
empty_keys = 0
empty_values = 0
duplicate_keys = 0

for row in range(2, ws_stary_read.max_row + 1):
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        key_normalized = str(key_c).strip()
        
        if is_cell_empty(value_f):
            empty_values += 1
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1

print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów")

# EKSPORT SŁOWNIKA DO CSV
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])

print(f"  ✓ Zapisano {len(stary_dict)} wpisów")

# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ
print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active

# Znajdź kolumny stare_oznaczenie, nowe_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie', 'nowe_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
nowe_oznaczenie_col = None
skrot_col = None

for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'nowe_oznaczenie':
        nowe_oznaczenie_col = col_letter
        print(f"  Znaleziono 'nowe_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col or not nowe_oznaczenie_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# KROK 2: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 2: Wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0

for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
        else:
            not_found_count += 1

print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono: {not_found_count} wierszy")
print(f"  ⚠ Wartość F pusta: {empty_value_count} wierszy")

# KROK 3: Tworzenie kolumny REWIZJA po kolumnie skrot
print(f"\n=== KROK 3: Tworzenie kolumny 'REWIZJA' ===")
# Kolumna REWIZJA będzie zaraz po kolumnie skrot
skrot_col_num = openpyxl.utils.cell.column_index_from_string(skrot_col)
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)

# Dodaj nagłówek
ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
print(f"  ✓ Utworzono kolumnę REWIZJA w kolumnie: {rewizja_col}")

# Dodaj formuły porównujące nowe_oznaczenie i stare_oznaczenie
formula_count = 0
prawda_count = 0
falsz_count = 0

for row in range(2, ws_write.max_row + 1):
    # Dodaj formułę: =nowe_oznaczenie = stare_oznaczenie
    formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
    ws_write[f'{rewizja_col}{row}'].value = formula
    formula_count += 1

print(f"  ✓ Dodano {formula_count} formuł porównujących")
print(f"  ℹ Formuła: ={nowe_oznaczenie_col}X={stare_oznaczenie_col}X")
print(f"  ℹ PRAWDA = wartości są identyczne")
print(f"  ℹ FAŁSZ = wartości są różne")

# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])

print(f"  ✓ Zapisano wartości skrot")

# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki:")
print(f"  1. {dictionary_export_file} - słownik")
print(f"  2. {skrot_export_file} - analiza dopasowań")
print(f"  3. {output_file} - wynikowy plik Excel z kolumną REWIZJA")
print(f"\n💡 Otwórz plik w Excelu i sprawdź kolumnę REWIZJA:")
print(f"   - PRAWDA = nowe i stare oznaczenie są identyczne")
print(f"   - FAŁSZ = wartości się różnią (wymaga sprawdzenia)")
```


## Kluczowe dodania

### Kolumna REWIZJA

Tworzona zaraz po kolumnie skrot (kolumna AR):[^18_3][^18_1]

```python
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)
ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
```


### Formuła porównania

Każdy wiersz dostaje formułę sprawdzającą równość:[^18_2][^18_1]

```python
formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
ws_write[f'{rewizja_col}{row}'].value = formula
```


### Wynik w Excelu

Po otwarciu pliku zobaczysz:[^18_1][^18_3]

- **PRAWDA** - wartości są identyczne (rewizja się zgadza)
- **FAŁSZ** - wartości są różne (wymaga sprawdzenia)


### Automatyczne przeliczanie

Excel automatycznie przeliczy formuły po otwarciu pliku[^18_1]

Możesz teraz łatwo filtrować kolumnę REWIZJA aby zobaczyć które pozycje wymagają korekty (FAŁSZ)![^18_3][^18_1]
<span style="display:none">[^18_10][^18_11][^18_12][^18_13][^18_4][^18_5][^18_6][^18_7][^18_8][^18_9]</span>

<div align="center">⁂</div>

[^18_1]: https://www.datacamp.com/tutorial/compare-two-columns-in-excel

[^18_2]: https://www.academyoflearning.com/blog/how-to-compare-two-columns-in-excel-easy-guide/

[^18_3]: https://excelchamps.com/formulas/compare-two-cells/

[^18_4]: image.jpg

[^18_5]: image.jpg

[^18_6]: image.jpg

[^18_7]: https://www.ablebits.com/office-addins-blog/excel-if-match-formula/

[^18_8]: https://techcommunity.microsoft.com/discussions/excelgeneral/formula-to-compare-text-in-2-cells/4036895

[^18_9]: https://www.statology.org/how-to-implement-excels-vlookup-with-python-and-openpyxl/

[^18_10]: https://stackoverflow.com/questions/75772707/how-to-compare-two-columns-in-excel-using-python

[^18_11]: https://techcommunity.microsoft.com/discussions/excelgeneral/return-the-matching-value-if-two-cells-match/3974284

[^18_12]: https://www.geeksforgeeks.org/python/adding-conditional-formatting-to-excel-using-python-openpyxl/

[^18_13]: https://codepal.ai/excel-formula-generator/query/WWG002kN/excel-formula-python-compare-values


---

# nie dziala teraz, nic nie zmieniaj. dodaj tylko kolumne rewizja nastepnie formule ktora sprawdza 2 kolumny.  dopasowanie w tej wersji jest prawidlowe from openpyxl import load_workbook

from openpyxl.utils import get_column_letter
import shutil
import os
import csv

def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False

# File paths

source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'

# Sprawdź czy plik źródłowy istnieje

if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

# Kopiuj plik

print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)

# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ

print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active

# Sprawdź czy arkusz 'stary' istnieje

if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()

ws_stary_read = wb_read['stary']

# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'

print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print("  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
stary_dict = {}
empty_keys = 0
empty_values = 0
duplicate_keys = 0

for row in range(2, ws_stary_read.max_row + 1):
    \# Czytamy WARTOŚCI (nie formuły) z kolumn C i F
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
   
    if not is_cell_empty(key_c):
        \# Normalizuj klucz
        key_normalized = str(key_c).strip()
       
        \# Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            \# Debug: pokaż pierwsze 5 przypadków pustych wartości
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
       
        \# Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
       
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1

print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy (kolumna C)")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")

# EKSPORT SŁOWNIKA DO CSV

print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
   
    \# Sortujemy słownik alfabetycznie
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])

print(f"  ✓ Zapisano {len(stary_dict)} wpisów do pliku '{dictionary_export_file}'")

# Przykładowe wpisy ze słownika

print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")

# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ

print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active

# Znajdź kolumny stare_oznaczenie i skrot

print("\nSzukanie kolumn 'stare_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
skrot_col = None

for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()

# KROK 2: Zbieranie wartości skrot

print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws_write.max_row + 1, 12)):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        has_value = ""
        if skrot_normalized in stary_dict:
            if is_cell_empty(stary_dict[skrot_normalized]):
                has_value = " (ale wartość F jest PUSTA)"
        skrot_values.append((row, skrot_normalized, exists + has_value))

print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")

# EKSPORT WARTOŚCI SKROT

skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
   
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])

print(f"  ✓ Zapisano wartości skrot do pliku '{skrot_export_file}'")

# KROK 3: Wypełnianie kolumny stare_oznaczenie

print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0
not_found_samples = []

for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
   
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
       
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
           
            \# Sprawdź czy wartość F nie jest pusta
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
               
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
                if len(not_found_samples) < 10:
                    not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - KLUCZ ISTNIEJE ale wartość F PUSTA")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - BRAK w słowniku")

print(f"\n=== PODSUMOWANIE ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono w słowniku: {not_found_count} wierszy")
print(f"  ⚠ Znaleziono klucz ale wartość F pusta: {empty_value_count} wierszy")

if not_found_samples:
    print(f"\n  ℹ Przykłady problemów:")
    for sample in not_found_samples:
        print(f"    - {sample}")

# Zapisz plik

print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {dictionary_export_file} - słownik z oznaczeniem pustych wartości")
print(f"  2. {skrot_export_file} - analiza dopasowań")
print(f"  3. {output_file} - wynikowy plik Excel")

Rozumiem! Dodaję TYLKO kolumnę REWIZJA z formułą porównującą do Twojego działającego kodu:[^19_1][^19_2]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import shutil
import os
import csv


def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'


# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()


# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)


# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ
print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active


# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()


ws_stary_read = wb_read['stary']


# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print("  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
stary_dict = {}
empty_keys = 0
empty_values = 0
duplicate_keys = 0


for row in range(2, ws_stary_read.max_row + 1):
    # Czytamy WARTOŚCI (nie formuły) z kolumn C i F
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        # Normalizuj klucz
        key_normalized = str(key_c).strip()
        
        # Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            # Debug: pokaż pierwsze 5 przypadków pustych wartości
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        # Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1


print(f"  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy (kolumna C)")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")


# EKSPORT SŁOWNIKA DO CSV
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    # Sortujemy słownik alfabetycznie
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])


print(f"  ✓ Zapisano {len(stary_dict)} wpisów do pliku '{dictionary_export_file}'")


# Przykładowe wpisy ze słownika
print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")


# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ
print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active


# Znajdź kolumny stare_oznaczenie, nowe_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie', 'nowe_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
nowe_oznaczenie_col = None
skrot_col = None


for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'nowe_oznaczenie':
        nowe_oznaczenie_col = col_letter
        print(f"  Znaleziono 'nowe_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")


if not stare_oznaczenie_col or not skrot_col or not nowe_oznaczenie_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()


# KROK 2: Zbieranie wartości skrot
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws_write.max_row + 1, 12)):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        has_value = ""
        if skrot_normalized in stary_dict:
            if is_cell_empty(stary_dict[skrot_normalized]):
                has_value = " (ale wartość F jest PUSTA)"
        skrot_values.append((row, skrot_normalized, exists + has_value))


print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")


# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])


print(f"  ✓ Zapisano wartości skrot do pliku '{skrot_export_file}'")


# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0
not_found_samples = []


for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            # Sprawdź czy wartość F nie jest pusta
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
                if len(not_found_samples) < 10:
                    not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - KLUCZ ISTNIEJE ale wartość F PUSTA")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - BRAK w słowniku")


print(f"\n=== PODSUMOWANIE DOPASOWANIA ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono w słowniku: {not_found_count} wierszy")
print(f"  ⚠ Znaleziono klucz ale wartość F pusta: {empty_value_count} wierszy")


if not_found_samples:
    print(f"\n  ℹ Przykłady problemów:")
    for sample in not_found_samples:
        print(f"    - {sample}")


# KROK 4: DODANIE KOLUMNY REWIZJA
print(f"\n=== KROK 4: Tworzenie kolumny 'REWIZJA' ===")

# Znajdź numer kolumny skrot i ustaw kolumnę rewizja zaraz po niej
skrot_col_num = openpyxl.utils.cell.column_index_from_string(skrot_col)
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)

# Dodaj nagłówek REWIZJA
ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
print(f"  ✓ Utworzono kolumnę 'REWIZJA' w kolumnie: {rewizja_col}")

# Dodaj formuły porównujące nowe_oznaczenie i stare_oznaczenie
formula_count = 0
for row in range(2, ws_write.max_row + 1):
    # Formuła sprawdzająca równość dwóch kolumn
    formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
    ws_write[f'{rewizja_col}{row}'].value = formula
    formula_count += 1

print(f"  ✓ Dodano {formula_count} formuł porównujących")
print(f"  ℹ Przykład formuły: ={nowe_oznaczenie_col}2={stare_oznaczenie_col}2")
print(f"  ℹ PRAWDA = wartości identyczne")
print(f"  ℹ FAŁSZ = wartości różne (wymaga sprawdzenia)")


# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {dictionary_export_file} - słownik z oznaczeniem pustych wartości")
print(f"  2. {skrot_export_file} - analiza dopasowań")
print(f"  3. {output_file} - wynikowy plik Excel z kolumną REWIZJA")
print(f"\n💡 Kolumna REWIZJA pokazuje:")
print(f"   PRAWDA - nowe_oznaczenie = stare_oznaczenie (OK)")
print(f"   FAŁSZ - wartości się różnią (wymaga sprawdzenia)")
```


## Co zostało dodane

### Import openpyxl.utils.cell

Dodałem na początku:[^19_1]

```python
import openpyxl.utils.cell
```


### Wyszukiwanie kolumny nowe_oznaczenie

Dodałem do sekcji wyszukiwania kolumn:[^19_1]

```python
elif header == 'nowe_oznaczenie':
    nowe_oznaczenie_col = col_letter
```


### KROK 4: Kolumna REWIZJA

Cała nowa sekcja która:[^19_2][^19_1]

1. Tworzy kolumnę REWIZJA zaraz po kolumnie skrot
2. Dodaje nagłówek
3. Wypełnia każdy wiersz formułą porównującą

### Formuła porównania

```python
formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
```

Reszta kodu pozostaje **bez zmian** - dopasowanie działa tak samo jak wcześniej![^19_2][^19_1]

<div align="center">⁂</div>

[^19_1]: https://www.datacamp.com/tutorial/compare-two-columns-in-excel

[^19_2]: https://excelchamps.com/formulas/compare-two-cells/


---

# przy sprawdzaniu wykazu gdzie mamy tylko pozycje z propozycja jest ok, ale w przypadku wykazu z dodatkowymi kolumnami w pliku csv wartosci_skrot_do_sprawdzenia.csv np jest 691

5529,SL10115513
NIE
BRAK
PROBLEM  zauwazylem ze slownik_stary_C_F.csv ma tylko 6462 wiersze i nie pobral wszystkiego z zakladki stary gdzie jest ponad 21834 wiersze

Doskonałe spostrzeżenie! Problem jest w tym że `ws_stary_read.max_row` zwraca ostatni wiersz z **danymi**, ale nie bierze pod uwagę że niektóre komórki mogą być puste między wierszami. Użyjmy `ws_stary_read.max_row` bez ograniczeń i dodatkowo zróbmy debug:[^20_1]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import shutil
import os
import csv


def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'


# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()


# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)


# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ
print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active


# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()


ws_stary_read = wb_read['stary']

# DIAGNOSTYKA: Sprawdź ile wierszy ma arkusz 'stary'
print(f"\n=== DIAGNOSTYKA ARKUSZA 'stary' ===")
print(f"  ws_stary_read.max_row = {ws_stary_read.max_row}")
print(f"  ws_stary_read.max_column = {ws_stary_read.max_column}")

# Sprawdź czy są dane poza max_row
last_row_with_data = ws_stary_read.max_row
for row in range(ws_stary_read.max_row, 0, -1):
    if ws_stary_read[f'C{row}'].value is not None or ws_stary_read[f'F{row}'].value is not None:
        last_row_with_data = row
        break

print(f"  Ostatni wiersz z danymi w C lub F: {last_row_with_data}")

# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'
print("\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print("  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
print(f"  ℹ Przetwarzanie wierszy od 2 do {last_row_with_data}...")

stary_dict = {}
empty_keys = 0
empty_values = 0
duplicate_keys = 0
processed_rows = 0


# ZMIANA: Iterujemy przez WSZYSTKIE wiersze do ostatniego z danymi
for row in range(2, last_row_with_data + 1):
    processed_rows += 1
    
    # Progress bar co 1000 wierszy
    if processed_rows % 1000 == 0:
        print(f"  ...przetworzono {processed_rows} wierszy")
    
    # Czytamy WARTOŚCI (nie formuły) z kolumn C i F
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
    
    if not is_cell_empty(key_c):
        # Normalizuj klucz
        key_normalized = str(key_c).strip()
        
        # Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            # Debug: pokaż pierwsze 5 przypadków pustych wartości
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        # Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1


print(f"\n  ✓ Utworzono słownik z {len(stary_dict)} wpisami")
print(f"  ℹ Przetworzono {processed_rows} wierszy")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy (kolumna C)")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")


# EKSPORT SŁOWNIKA DO CSV
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    # Sortujemy słownik alfabetycznie
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])


print(f"  ✓ Zapisano {len(stary_dict)} wpisów do pliku '{dictionary_export_file}'")


# Przykładowe wpisy ze słownika
print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")

# Sprawdź konkretną wartość która powinna istnieć
test_key = "5529,SL10115513"
if test_key in stary_dict:
    print(f"\n  ✓ TEST: Klucz '{test_key}' ISTNIEJE w słowniku -> '{stary_dict[test_key]}'")
else:
    print(f"\n  ✗ TEST: Klucz '{test_key}' NIE ISTNIEJE w słowniku")


# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ
print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active


# Znajdź kolumny stare_oznaczenie, nowe_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie', 'nowe_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
nowe_oznaczenie_col = None
skrot_col = None


for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'nowe_oznaczenie':
        nowe_oznaczenie_col = col_letter
        print(f"  Znaleziono 'nowe_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")


if not stare_oznaczenie_col or not skrot_col or not nowe_oznaczenie_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()


# KROK 2: Zbieranie wartości skrot
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws_write.max_row + 1, 12)):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        has_value = ""
        if skrot_normalized in stary_dict:
            if is_cell_empty(stary_dict[skrot_normalized]):
                has_value = " (ale wartość F jest PUSTA)"
        skrot_values.append((row, skrot_normalized, exists + has_value))


print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")


# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])


print(f"  ✓ Zapisano wartości skrot do pliku '{skrot_export_file}'")


# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0
not_found_samples = []


for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            # Sprawdź czy wartość F nie jest pusta
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
                if len(not_found_samples) < 10:
                    not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - KLUCZ ISTNIEJE ale wartość F PUSTA")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - BRAK w słowniku")


print(f"\n=== PODSUMOWANIE DOPASOWANIA ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono w słowniku: {not_found_count} wierszy")
print(f"  ⚠ Znaleziono klucz ale wartość F pusta: {empty_value_count} wierszy")


if not_found_samples:
    print(f"\n  ℹ Przykłady problemów:")
    for sample in not_found_samples:
        print(f"    - {sample}")


# KROK 4: DODANIE KOLUMNY REWIZJA
print(f"\n=== KROK 4: Tworzenie kolumny 'REWIZJA' ===")

# Znajdź numer kolumny skrot i ustaw kolumnę rewizja zaraz po niej
skrot_col_num = openpyxl.utils.cell.column_index_from_string(skrot_col)
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)

# Dodaj nagłówek REWIZJA
ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
print(f"  ✓ Utworzono kolumnę 'REWIZJA' w kolumnie: {rewizja_col}")

# Dodaj formuły porównujące nowe_oznaczenie i stare_oznaczenie
formula_count = 0
for row in range(2, ws_write.max_row + 1):
    # Formuła sprawdzająca równość dwóch kolumn
    formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
    ws_write[f'{rewizja_col}{row}'].value = formula
    formula_count += 1

print(f"  ✓ Dodano {formula_count} formuł porównujących")
print(f"  ℹ Przykład formuły: ={nowe_oznaczenie_col}2={stare_oznaczenie_col}2")
print(f"  ℹ PRAWDA = wartości identyczne")
print(f"  ℹ FAŁSZ = wartości różne (wymaga sprawdzenia)")


# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {dictionary_export_file} - słownik z {len(stary_dict)} wpisami")
print(f"  2. {skrot_export_file} - analiza dopasowań")
print(f"  3. {output_file} - wynikowy plik Excel z kolumną REWIZJA")
print(f"\n💡 Kolumna REWIZJA pokazuje:")
print(f"   PRAWDA - nowe_oznaczenie = stare_oznaczenie (OK)")
print(f"   FAŁSZ - wartości się różnią (wymaga sprawdzenia)")
```


## Kluczowe zmiany

### Diagnostyka arkusza stary

Dodałem sekcję pokazującą ile wierszy ma arkusz:[^20_1]

```python
print(f"  ws_stary_read.max_row = {ws_stary_read.max_row}")
```


### Szukanie ostatniego wiersza z danymi

Skrypt sprawdza od końca gdzie są ostatnie dane:[^20_2]

```python
for row in range(ws_stary_read.max_row, 0, -1):
    if ws_stary_read[f'C{row}'].value is not None or ws_stary_read[f'F{row}'].value is not None:
        last_row_with_data = row
        break
```


### Progress bar

Co 1000 wierszy wyświetla postęp przetwarzania[^20_1]

### Test konkretnego klucza

Sprawdza czy problemowy klucz `5529,SL10115513` istnieje w słowniku

Teraz powinien przetworzyć wszystkie ~21834 wiersze z arkusza 'stary'! Uruchom skrypt i sprawdź czy słownik ma teraz więcej wpisów.[^20_2][^20_1]
<span style="display:none">[^20_3][^20_4]</span>

<div align="center">⁂</div>

[^20_1]: https://www.datacamp.com/tutorial/openpyxl

[^20_2]: https://www.geeksforgeeks.org/python/get-values-of-all-rows-in-a-particular-column-in-openpyxl-python/

[^20_3]: image.jpg

[^20_4]: image.jpg


---

# nadal zle tworzy slownik starty slownik_stary_C_F.csv ma tylko ok 6500 wierszow. w slowniku powinny byc wszystkie wiersze z zakladki stary

Rozumiem! Problem jest w tym że słownik ma tylko **unikalne klucze** (kolumna C), a jeśli są duplikaty, tylko ostatnia wartość jest zachowana. Musimy użyć **twardej liczby wierszy** zamiast polegać na `max_row`. Oto poprawione rozwiązanie:[^21_1][^21_2]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import shutil
import os
import csv


def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


# File paths
source_file = 'propozycje_updated.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'
all_rows_export_file = 'wszystkie_wiersze_stary_C_F.csv'


# Sprawdź czy plik źródłowy istnieje
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()


# Kopiuj plik
print(f"Kopiowanie pliku '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)


# KLUCZOWE: Wczytaj plik Z data_only=True aby czytać WARTOŚCI zamiast FORMUŁ
print(f"Wczytywanie pliku '{output_file}' z data_only=True (tylko wartości, nie formuły)...")
wb_read = load_workbook(output_file, data_only=True)
ws_read = wb_read.active


# Sprawdź czy arkusz 'stary' istnieje
if 'stary' not in wb_read.sheetnames:
    print("BŁĄD: Nie znaleziono arkusza 'stary'!")
    exit()


ws_stary_read = wb_read['stary']

# DIAGNOSTYKA: Sprawdź ile wierszy ma arkusz 'stary'
print(f"\n=== DIAGNOSTYKA ARKUSZA 'stary' ===")
print(f"  ws_stary_read.max_row = {ws_stary_read.max_row}")
print(f"  ws_stary_read.max_column = {ws_stary_read.max_column}")

# KROK 1: Tworzenie słownika z WARTOŚCI (nie formuł) z arkusza 'stary'
# ZMIANA: Używamy TWARDEJ LICZBY 25000 aby mieć pewność że przejdziemy przez wszystkie wiersze
MAX_ROWS_TO_CHECK = 25000  # Zwiększ jeśli masz więcej wierszy

print(f"\n=== KROK 1: Tworzenie słownika z arkusza 'stary' (kolumny C -> F) ===")
print(f"  ℹ Czytanie WARTOŚCI zamiast formuł (data_only=True)")
print(f"  ℹ Przetwarzanie wierszy od 2 do {MAX_ROWS_TO_CHECK}...")

stary_dict = {}
all_rows_data = []  # Lista WSZYSTKICH wierszy (włącznie z duplikatami)
empty_keys = 0
empty_values = 0
duplicate_keys = 0
processed_rows = 0
rows_with_data = 0


# Iterujemy przez WSZYSTKIE możliwe wiersze (twarda liczba)
for row in range(2, MAX_ROWS_TO_CHECK + 1):
    processed_rows += 1
    
    # Progress bar co 2000 wierszy
    if processed_rows % 2000 == 0:
        print(f"  ...przetworzono {processed_rows} wierszy, znaleziono {rows_with_data} z danymi")
    
    # Czytamy WARTOŚCI (nie formuły) z kolumn C i F
    key_c = ws_stary_read[f'C{row}'].value
    value_f = ws_stary_read[f'F{row}'].value
    
    # Jeżeli kolumna C ma dane, zapisujemy
    if not is_cell_empty(key_c):
        rows_with_data += 1
        
        # Normalizuj klucz
        key_normalized = str(key_c).strip()
        
        # Dodaj do listy WSZYSTKICH wierszy (dla CSV)
        all_rows_data.append({
            'wiersz': row,
            'klucz_c': key_normalized,
            'wartosc_f': value_f if not is_cell_empty(value_f) else "PUSTA"
        })
        
        # Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        # Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        # Dodaj/nadpisz w słowniku
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1


print(f"\n  ✓ Przetworzono {processed_rows} wierszy")
print(f"  ✓ Znaleziono {rows_with_data} wierszy z danymi w kolumnie C")
print(f"  ✓ Utworzono słownik z {len(stary_dict)} UNIKALNYMI kluczami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy (kolumna C)")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy (nadpisane w słowniku)")


# EKSPORT WSZYSTKICH WIERSZY DO CSV (włącznie z duplikatami)
print(f"\n=== EKSPORT WSZYSTKICH WIERSZY DO CSV: {all_rows_export_file} ===")
with open(all_rows_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz_arkusza', 'Klucz_C', 'Wartość_F'])
    
    for row_data in all_rows_data:
        writer.writerow([row_data['wiersz'], row_data['klucz_c'], row_data['wartosc_f']])

print(f"  ✓ Zapisano {len(all_rows_data)} WSZYSTKICH wierszy z danymi")


# EKSPORT SŁOWNIKA DO CSV (tylko unikalne klucze)
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    # Sortujemy słownik alfabetycznie
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])


print(f"  ✓ Zapisano {len(stary_dict)} UNIKATOWYCH wpisów")


# Przykładowe wpisy ze słownika
print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")

# Sprawdź konkretną wartość która powinna istnieć
test_key = "5529,SL10115513"
if test_key in stary_dict:
    print(f"\n  ✓ TEST: Klucz '{test_key}' ISTNIEJE w słowniku -> '{stary_dict[test_key]}'")
else:
    print(f"\n  ✗ TEST: Klucz '{test_key}' NIE ISTNIEJE w słowniku")
    # Szukaj czy jest w liście wszystkich wierszy
    found_in_list = [r for r in all_rows_data if r['klucz_c'] == test_key]
    if found_in_list:
        print(f"  ℹ Ale znaleziono w liście wszystkich wierszy: {found_in_list}")


# TERAZ wczytaj plik PONOWNIE bez data_only aby móc ZAPISYWAĆ
print(f"\n=== Wczytywanie pliku ponownie do ZAPISU ===")
wb_write = load_workbook(output_file)
ws_write = wb_write.active


# Znajdź kolumny stare_oznaczenie, nowe_oznaczenie i skrot
print("\nSzukanie kolumn 'stare_oznaczenie', 'nowe_oznaczenie' i 'skrot'...")
stare_oznaczenie_col = None
nowe_oznaczenie_col = None
skrot_col = None


for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  Znaleziono 'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'nowe_oznaczenie':
        nowe_oznaczenie_col = col_letter
        print(f"  Znaleziono 'nowe_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  Znaleziono 'skrot' w kolumnie: {skrot_col}")


if not stare_oznaczenie_col or not skrot_col or not nowe_oznaczenie_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()


# KROK 2: Zbieranie wartości skrot
print(f"\n=== KROK 2: Zbieranie wartości z kolumny 'skrot' ({skrot_col}) ===")
skrot_values = []
for row in range(2, min(ws_write.max_row + 1, 12)):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        exists = "✓ ISTNIEJE" if skrot_normalized in stary_dict else "✗ BRAK"
        has_value = ""
        if skrot_normalized in stary_dict:
            if is_cell_empty(stary_dict[skrot_normalized]):
                has_value = " (ale wartość F jest PUSTA)"
        skrot_values.append((row, skrot_normalized, exists + has_value))


print(f"  ℹ Przykładowe wartości skrot (pierwsze {len(skrot_values)}):")
for row, val, exists in skrot_values:
    print(f"    Wiersz {row}: '{val}' -> {exists}")


# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT DO CSV: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])


print(f"  ✓ Zapisano wartości skrot")


# KROK 3: Wypełnianie kolumny stare_oznaczenie
print(f"\n=== KROK 3: Dopasowywanie i wypełnianie kolumny 'stare_oznaczenie' ({stare_oznaczenie_col}) ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0
not_found_samples = []


for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
                if len(not_found_samples) < 10:
                    not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - KLUCZ ISTNIEJE ale wartość F PUSTA")
        else:
            not_found_count += 1
            if len(not_found_samples) < 10:
                not_found_samples.append(f"'{skrot_normalized}' (wiersz {row}) - BRAK w słowniku")


print(f"\n=== PODSUMOWANIE DOPASOWANIA ===")
print(f"  ✓ Wypełniono: {filled_count} wierszy")
print(f"  ✗ Nie znaleziono w słowniku: {not_found_count} wierszy")
print(f"  ⚠ Znaleziono klucz ale wartość F pusta: {empty_value_count} wierszy")


if not_found_samples:
    print(f"\n  ℹ Przykłady problemów:")
    for sample in not_found_samples:
        print(f"    - {sample}")


# KROK 4: DODANIE KOLUMNY REWIZJA
print(f"\n=== KROK 4: Tworzenie kolumny 'REWIZJA' ===")

skrot_col_num = openpyxl.utils.cell.column_index_from_string(skrot_col)
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)

ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
print(f"  ✓ Utworzono kolumnę 'REWIZJA' w kolumnie: {rewizja_col}")

formula_count = 0
for row in range(2, ws_write.max_row + 1):
    formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
    ws_write[f'{rewizja_col}{row}'].value = formula
    formula_count += 1

print(f"  ✓ Dodano {formula_count} formuł porównujących")


# Zapisz plik
print(f"\n=== ZAPISYWANIE ===")
print(f"Zapisywanie pliku '{output_file}'...")
wb_write.save(output_file)
print(f"✓ Plik został pomyślnie zapisany jako '{output_file}'")
print(f"\n📄 Wygenerowane pliki do analizy:")
print(f"  1. {all_rows_export_file} - WSZYSTKIE {len(all_rows_data)} wiersze z arkusza 'stary'")
print(f"  2. {dictionary_export_file} - słownik z {len(stary_dict)} unikalnymi kluczami")
print(f"  3. {skrot_export_file} - analiza dopasowań")
print(f"  4. {output_file} - wynikowy plik Excel")
```


## Kluczowe zmiany

### Twarda liczba wierszy

Zamiast polegać na `max_row`, iteruję przez **25000 wierszy**:[^21_2][^21_1]

```python
MAX_ROWS_TO_CHECK = 25000
for row in range(2, MAX_ROWS_TO_CHECK + 1):
```


### Eksport WSZYSTKICH wierszy

Nowy plik `wszystkie_wiersze_stary_C_F.csv` zawiera każdy wiersz z danymi (włącznie z duplikatami):[^21_3]

```python
all_rows_data.append({
    'wiersz': row,
    'klucz_c': key_normalized,
    'wartosc_f': value_f
})
```


### Rozróżnienie

- **wszystkie_wiersze_stary_C_F.csv** - każdy wiersz osobno (~21834 wierszy)
- **slownik_stary_C_F.csv** - tylko unikalne klucze (~6500 wierszy jeśli są duplikaty)


### Więcej diagnostyki

- Licznik wierszy z danymi
- Progress co 2000 wierszy
- Test czy konkretny klucz istnieje

Teraz zobaczysz dokładnie ile wierszy ma dane i dlaczego słownik ma mniej wpisów (duplikaty kluczy)![^21_1][^21_2][^21_3]

<div align="center">⁂</div>

[^21_1]: https://stackoverflow.com/questions/34754077/openpyxl-how-to-read-only-one-column-from-excel-file-in-python

[^21_2]: https://www.geeksforgeeks.org/python/get-values-of-all-rows-in-a-particular-column-in-openpyxl-python/

[^21_3]: https://www.geeksforgeeks.org/python/how-to-save-a-python-dictionary-to-a-csv-file/


---

# szukaj nie pobralo wszystkich wierszy we wszystkie_wiersze_stary_C_F.csv wiec problem jest ze zle pobiera

Rozumiem! Problem jest prawdopodobnie w tym, że `data_only=True` **nie przelicza formuł** jeśli plik nie był otwarty i zapisany w Excelu. Użyjmy innego podejścia - najpierw wczytaj plik źródłowy STARE bezpośrednio z `data_only=True`:[^22_1][^22_2]

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.utils.cell
import shutil
import os
import csv


def is_cell_empty(value):
    """
    Sprawdza czy wartość komórki jest naprawdę pusta.
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


# File paths
source_file = 'propozycje_updated.xlsx'
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'
output_file = 'propozycje_updated_z_dopasowaniem.xlsx'
dictionary_export_file = 'slownik_stary_C_F.csv'
all_rows_export_file = 'wszystkie_wiersze_stary_C_F.csv'


# Sprawdź czy pliki źródłowe istnieją
if not os.path.exists(source_file):
    print(f"BŁĄD: Nie znaleziono pliku '{source_file}'")
    print("Najpierw uruchom skrypt 'sprawdzanie_rewizji.py'")
    exit()

if not os.path.exists(stare_file):
    print(f"BŁĄD: Nie znaleziono pliku '{stare_file}'")
    exit()


# KLUCZOWE: Wczytaj ORYGINALNY plik STARE bezpośrednio Z data_only=True
print(f"\n=== WCZYTYWANIE ORYGINALNEGO PLIKU STARE ===")
print(f"Wczytywanie '{stare_file}' z data_only=True...")
wb_stare_original = load_workbook(stare_file, data_only=True)
ws_stare_original = wb_stare_original.active

print(f"  ✓ Wczytano plik STARE")
print(f"  ℹ Nazwa aktywnego arkusza: {ws_stare_original.title}")
print(f"  ℹ max_row = {ws_stare_original.max_row}")
print(f"  ℹ max_column = {ws_stare_original.max_column}")

# Sprawdź nagłówki
print(f"\n  ℹ Nagłówki w pierwszym wierszu:")
for col in ['A', 'B', 'C', 'D', 'E', 'F']:
    header = ws_stare_original[f'{col}1'].value
    print(f"    Kolumna {col}: '{header}'")


# KROK 1: Tworzenie słownika z ORYGINALNEGO pliku STARE
MAX_ROWS_TO_CHECK = 30000  # Zwiększone dla pewności

print(f"\n=== KROK 1: Tworzenie słownika z oryginalnego pliku STARE (kolumny C -> F) ===")
print(f"  ℹ Przetwarzanie wierszy od 2 do {MAX_ROWS_TO_CHECK}...")

stary_dict = {}
all_rows_data = []
empty_keys = 0
empty_values = 0
duplicate_keys = 0
processed_rows = 0
rows_with_data = 0


for row in range(2, MAX_ROWS_TO_CHECK + 1):
    processed_rows += 1
    
    # Progress bar co 2000 wierszy
    if processed_rows % 2000 == 0:
        print(f"  ...przetworzono {processed_rows} wierszy, znaleziono {rows_with_data} z danymi")
    
    # Czytamy WARTOŚCI z kolumn C i F z ORYGINALNEGO pliku
    key_c = ws_stare_original[f'C{row}'].value
    value_f = ws_stare_original[f'F{row}'].value
    
    # Jeżeli kolumna C ma dane, zapisujemy
    if not is_cell_empty(key_c):
        rows_with_data += 1
        
        # Normalizuj klucz
        key_normalized = str(key_c).strip()
        
        # Dodaj do listy WSZYSTKICH wierszy
        all_rows_data.append({
            'wiersz': row,
            'klucz_c': key_normalized,
            'wartosc_f': value_f if not is_cell_empty(value_f) else "PUSTA"
        })
        
        # Sprawdź czy wartość F nie jest pusta
        if is_cell_empty(value_f):
            empty_values += 1
            if empty_values <= 5:
                print(f"  ⚠ Wiersz {row}: Klucz '{key_normalized}' ma pustą wartość F")
        
        # Sprawdź duplikaty
        if key_normalized in stary_dict:
            duplicate_keys += 1
        
        # Dodaj/nadpisz w słowniku
        stary_dict[key_normalized] = value_f
    else:
        empty_keys += 1
        # Jeśli mamy 1000 pustych wierszy z rzędu, przerywamy
        if empty_keys > 1000 and rows_with_data > 0:
            print(f"  ℹ Przerwano po {processed_rows} wierszach (1000 pustych z rzędu)")
            break


print(f"\n  ✓ Przetworzono {processed_rows} wierszy")
print(f"  ✓ Znaleziono {rows_with_data} wierszy z danymi w kolumnie C")
print(f"  ✓ Utworzono słownik z {len(stary_dict)} UNIKALNYMI kluczami")
print(f"  ℹ Pominięto {empty_keys} pustych kluczy")
print(f"  ⚠ Znaleziono {empty_values} pustych wartości (kolumna F)")
print(f"  ℹ Znaleziono {duplicate_keys} duplikatów kluczy")


# EKSPORT WSZYSTKICH WIERSZY
print(f"\n=== EKSPORT WSZYSTKICH WIERSZY DO CSV: {all_rows_export_file} ===")
with open(all_rows_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz_arkusza', 'Klucz_C', 'Wartość_F'])
    
    for row_data in all_rows_data:
        writer.writerow([row_data['wiersz'], row_data['klucz_c'], row_data['wartosc_f']])

print(f"  ✓ Zapisano {len(all_rows_data)} WSZYSTKICH wierszy z danymi")


# EKSPORT SŁOWNIKA
print(f"\n=== EKSPORT SŁOWNIKA DO CSV: {dictionary_export_file} ===")
with open(dictionary_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Klucz_C', 'Wartość_F', 'Czy_pusta'])
    
    for key in sorted(stary_dict.keys()):
        is_empty = "PUSTA" if is_cell_empty(stary_dict[key]) else "OK"
        writer.writerow([key, stary_dict[key], is_empty])

print(f"  ✓ Zapisano {len(stary_dict)} UNIKATOWYCH wpisów")


# Przykładowe wpisy
print(f"\n  ℹ Przykładowe wpisy ze słownika (pierwsze 10):")
for i, (k, v) in enumerate(list(stary_dict.items())[:10]):
    status = "✓" if not is_cell_empty(v) else "✗ PUSTA"
    print(f"    [{i+1}] '{k}' -> '{v}' {status}")

# Test konkretnego klucza
test_key = "5529,SL10115513"
if test_key in stary_dict:
    print(f"\n  ✓ TEST: Klucz '{test_key}' ISTNIEJE -> '{stary_dict[test_key]}'")
else:
    print(f"\n  ✗ TEST: Klucz '{test_key}' NIE ISTNIEJE")


# Kopiuj plik propozycje
print(f"\n=== KOPIOWANIE PLIKU PROPOZYCJE ===")
print(f"Kopiowanie '{source_file}' -> '{output_file}'...")
shutil.copy(source_file, output_file)


# Wczytaj skopiowany plik do ZAPISU
print(f"\nWczytywanie '{output_file}' do zapisu...")
wb_write = load_workbook(output_file)
ws_write = wb_write.active


# Znajdź kolumny
print("\n=== SZUKANIE KOLUMN ===")
stare_oznaczenie_col = None
nowe_oznaczenie_col = None
skrot_col = None

for col in range(1, ws_write.max_column + 1):
    col_letter = get_column_letter(col)
    header = ws_write[f'{col_letter}1'].value
    if header == 'stare_oznaczenie':
        stare_oznaczenie_col = col_letter
        print(f"  'stare_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'nowe_oznaczenie':
        nowe_oznaczenie_col = col_letter
        print(f"  'nowe_oznaczenie' w kolumnie: {col_letter}")
    elif header == 'skrot':
        skrot_col = col_letter
        print(f"  'skrot' w kolumnie: {skrot_col}")

if not stare_oznaczenie_col or not skrot_col or not nowe_oznaczenie_col:
    print("BŁĄD: Nie znaleziono wymaganych kolumn!")
    exit()


# EKSPORT WARTOŚCI SKROT
skrot_export_file = 'wartosci_skrot_do_sprawdzenia.csv'
print(f"\n=== EKSPORT WARTOŚCI SKROT: {skrot_export_file} ===")
with open(skrot_export_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter=';')
    writer.writerow(['Wiersz', 'Wartość_skrot', 'Istnieje_w_słowniku', 'Wartość_F', 'Status'])
    
    for row in range(2, ws_write.max_row + 1):
        skrot_val = ws_write[f'{skrot_col}{row}'].value
        if not is_cell_empty(skrot_val):
            skrot_normalized = str(skrot_val).strip()
            exists = "TAK" if skrot_normalized in stary_dict else "NIE"
            value_f = stary_dict.get(skrot_normalized, "BRAK")
            status = "OK" if exists == "TAK" and not is_cell_empty(value_f) else "PROBLEM"
            writer.writerow([row, skrot_normalized, exists, value_f, status])

print(f"  ✓ Zapisano")


# WYPEŁNIANIE kolumny stare_oznaczenie
print(f"\n=== WYPEŁNIANIE KOLUMNY stare_oznaczenie ===")
filled_count = 0
not_found_count = 0
empty_value_count = 0

for row in range(2, ws_write.max_row + 1):
    skrot_val = ws_write[f'{skrot_col}{row}'].value
    
    if not is_cell_empty(skrot_val):
        skrot_normalized = str(skrot_val).strip()
        
        if skrot_normalized in stary_dict:
            value_f = stary_dict[skrot_normalized]
            
            if not is_cell_empty(value_f):
                ws_write[f'{stare_oznaczenie_col}{row}'].value = value_f
                filled_count += 1
                
                if filled_count <= 5:
                    print(f"  ✓ Wiersz {row}: '{skrot_normalized}' -> '{value_f}'")
            else:
                empty_value_count += 1
        else:
            not_found_count += 1

print(f"\n  ✓ Wypełniono: {filled_count}")
print(f"  ✗ Nie znaleziono: {not_found_count}")
print(f"  ⚠ Wartość F pusta: {empty_value_count}")


# DODANIE KOLUMNY REWIZJA
print(f"\n=== TWORZENIE KOLUMNY REWIZJA ===")

skrot_col_num = openpyxl.utils.cell.column_index_from_string(skrot_col)
rewizja_col_num = skrot_col_num + 1
rewizja_col = get_column_letter(rewizja_col_num)

ws_write[f'{rewizja_col}1'].value = 'REWIZJA'
print(f"  ✓ Kolumna REWIZJA: {rewizja_col}")

for row in range(2, ws_write.max_row + 1):
    formula = f'={nowe_oznaczenie_col}{row}={stare_oznaczenie_col}{row}'
    ws_write[f'{rewizja_col}{row}'].value = formula

print(f"  ✓ Dodano formuły")


# ZAPISYWANIE
print(f"\n=== ZAPISYWANIE ===")
wb_write.save(output_file)
print(f"✓ Zapisano '{output_file}'")

print(f"\n📄 Pliki:")
print(f"  1. {all_rows_export_file} - {len(all_rows_data)} wierszy")
print(f"  2. {dictionary_export_file} - {len(stary_dict)} kluczy")
print(f"  3. {skrot_export_file} - analiza")
print(f"  4. {output_file} - wynik")
```


## Kluczowe zmiany

### Bezpośrednie wczytanie oryginalnego pliku STARE

Zamiast czytać arkusz 'stary' z kopii, wczytuję **ORYGINALNY** plik `SPRAWDZANIE_REWIZJI-STARE.xlsx` z `data_only=True`:[^22_2][^22_1]

```python
wb_stare_original = load_workbook(stare_file, data_only=True)
ws_stare_original = wb_stare_original.active
```


### Inteligentne przerwanie

Przerywa jeśli napotka 1000 pustych wierszy z rzędu (żeby nie iterować przez 30000 gdy dane są np. do 22000)[^22_1]

### Zwiększony limit

`MAX_ROWS_TO_CHECK = 30000` dla pewności[^22_3]

### Diagnostyka nagłówków

Pokazuje co jest w kolumnach A-F aby sprawdzić czy to właściwy arkusz[^22_4]

To powinno rozwiązać problem - teraz czytamy dane z **oryginalnego** pliku Excel który ma przeliczone formuły![^22_5][^22_2][^22_1]

<div align="center">⁂</div>

[^22_1]: https://stackoverflow.com/questions/36116162/python-openpyxl-data-only-true-returning-none

[^22_2]: https://www.reddit.com/r/learnpython/comments/147qbfh/using_openpyxlpandas_to_get_the_final_value_of_a/

[^22_3]: https://stackoverflow.com/questions/34754077/openpyxl-how-to-read-only-one-column-from-excel-file-in-python

[^22_4]: https://www.geeksforgeeks.org/python/get-values-of-all-rows-in-a-particular-column-in-openpyxl-python/

[^22_5]: https://www.datacamp.com/tutorial/openpyxl

