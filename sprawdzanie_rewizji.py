import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

# File paths
propozycje_file = 'propozycje.xlsx'  # Assuming it's in the current directory or provide full path
stare_file = r'\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\SPRAWDZANIE REWIZJI\SPRAWDZANIE_REWIZJI-STARE.xlsx'

# Load workbooks
wb_propozycje = load_workbook(propozycje_file)
ws_propozycje = wb_propozycje.active

# Assuming Zeinr is column A, ZEF is column Z, ZEV is column AA (adjust if needed)
zeinr_col = 'A'  # Zeinr column
zef_col = 'Z'    # ZEF column
zev_col = 'AA'   # ZEV column

# Create ZEFZEV column next to ZEV
zefzev_col = get_column_letter(openpyxl.utils.column_index_from_string(zev_col) + 1)
ws_propozycje[f'{zefzev_col}1'].value = 'ZEFZEV'

# Fill ZEFZEV with ZEF&ZEV if both ZEF and ZEV are present
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
for col in range(1, 7):  # A to F
    col_letter = get_column_letter(col)
    for row in range(1, ws_stare_source.max_row + 1):
        ws_stary[f'{col_letter}{row}'].value = ws_stare_source[f'{col_letter}{row}'].value

# Task 4: Split 'propozycja' column by '_' and place into columns starting from BH
propozycja_col = 'BH'  # Assuming propozycja is column BH, adjust if needed
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
# The example is =BN5&","&BJ5, so BN is nowe_oznaczenie, BJ is the first part after split
bj_col = 'BJ'  # Assuming BJ is the column for first split part
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
            # Perform VLOOKUP manually
            for stary_row in range(2, ws_stary.max_row + 1):
                if ws_stary[f'C{stary_row}'].value == skrot_val:  # Assuming lookup in column C
                    ws_propozycje[f'{stare_oznaczenie_col}{row}'].value = ws_stary[f'F{stary_row}'].value  # Column F is 4th column from C
                    break

# Save the workbook
wb_propozycje.save('propozycje_updated.xlsx')
