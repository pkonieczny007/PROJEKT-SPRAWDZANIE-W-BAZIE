import pandas as pd
import re
import os
import shutil
from pathlib import Path

def utworz_wykaz():
    """
    Funkcja tworzy wykaz.xlsx z wybranego pliku Excel.
    Dodaje kolumnę 'Nazwa' = Zeinr_p_posn
    """
    print("\n=== TWORZENIE WYKAZU ===")
    print("Dostępne pliki Excel w bieżącym folderze:")
    
    # Pokaż dostępne pliki .xlsx i .xlsm (z wyłączeniem wykaz.xlsx i propozycje.xlsx)
    pliki = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xlsm'))
             and f not in ['wykaz.xlsx', 'propozycje.xlsx']]
    
    if not pliki:
        print("❌ Nie znaleziono żadnych plików Excel w folderze!")
        return False
    
    for i, plik in enumerate(pliki, 1):
        print(f"{i}. {plik}")
    
    # Wybór pliku
    try:
        wybor = int(input("\nWybierz numer pliku (lub 0 aby anulować): "))
        if wybor == 0:
            print("Anulowano.")
            return False
        if wybor < 1 or wybor > len(pliki):
            print("❌ Nieprawidłowy wybór!")
            return False
        
        plik_zrodlowy = pliki[wybor - 1]
    except ValueError:
        print("❌ Nieprawidłowa wartość!")
        return False
    
    # Wczytaj dane
    print(f"\n📂 Wczytuję dane z: {plik_zrodlowy}")
    df = pd.read_excel(plik_zrodlowy, dtype=str)
    
    # Sprawdź czy są kolumny Zeinr i Posn
    if 'Zeinr' not in df.columns or 'Posn' not in df.columns:
        print("❌ Błąd: Plik musi zawierać kolumny 'Zeinr' i 'Posn'!")
        print(f"Znalezione kolumny: {', '.join(df.columns)}")
        return False

    # Znajdź kolumnę 'zakupy' (bez uwzględniania wielkości liter)
    zakupy_col = None
    for col in df.columns:
        if col.lower() == 'zakupy':
            zakupy_col = col
            break

    if not zakupy_col:
        print("❌ Błąd: Nie znaleziono kolumny 'zakupy'!")
        print(f"Znalezione kolumny: {', '.join(df.columns)}")
        return False

    # Uzupełnij puste wartości
    df['Zeinr'] = df['Zeinr'].fillna('')
    df['Posn'] = df['Posn'].fillna('')
    df[zakupy_col] = df[zakupy_col].fillna('')

    # Utwórz kolumnę Nazwa tylko dla elementów z 'blacha' w kolumnie zakupy
    df['Nazwa'] = ''
    mask = df[zakupy_col].str.lower().str.contains('blacha', na=False)
    df.loc[mask, 'Nazwa'] = df.loc[mask, 'Zeinr'] + '_p' + df.loc[mask, 'Posn']
    
    # Dodaj pustą kolumnę propozycja
    df['propozycja'] = ''
    
    # Zapisz wykaz
    df.to_excel('wykaz.xlsx', index=False)
    print(f"✅ Utworzono wykaz.xlsx z {len(df)} wierszami")
    print(f"   Kolumny: {', '.join(df.columns)}")
    return True


def uruchom_propozycje():
    """
    Funkcja uruchamia dopasowanie propozycji z pliku elementy1.xlsx
    """
    print("\n=== DOPASOWANIE PROPOZYCJI ===")
    
    # Sprawdź czy istnieje wykaz.xlsx
    if not os.path.exists('wykaz.xlsx'):
        print("❌ Błąd: Nie znaleziono pliku wykaz.xlsx!")
        return False
    
    # Sprawdź czy istnieje elementy1.xlsx
    if not os.path.exists('elementy1.xlsx'):
        print("❌ Nie znaleziono pliku elementy1.xlsx w folderze lokalnym!")
        print("   Pobieram najnowszą wersję z bazy danych...")

        # Ścieżka do pliku źródłowego
        sciezka_zrodlowa = r"\\QNAP-ENERGO\Technologia\BAZA\POBIERANIE Z BAZY\elementy1.xlsx"

        if not os.path.exists(sciezka_zrodlowa):
            print("❌ Błąd: Nie można znaleźć pliku źródłowego w bazie danych!")
            return False

        try:
            # Skopiuj plik
            shutil.copy2(sciezka_zrodlowa, 'elementy1.xlsx')
            print("✅ Pobrano plik elementy1.xlsx z bazy danych")

            # Wczytaj datę utworzenia z komórki B2
            df_data = pd.read_excel('elementy1.xlsx', header=None)
            if df_data.shape[0] > 1 and df_data.shape[1] > 1:
                data_utworzenia = df_data.iloc[1, 1]  # B2
                print(f"📅 Data utworzenia bazy: {data_utworzenia}")
            else:
                print("⚠️  Nie udało się odczytać daty utworzenia z pliku")

        except Exception as e:
            print(f"❌ Błąd podczas pobierania pliku: {e}")
            return False
    
    # Wczytaj dane
    print("📂 Wczytuję dane...")
    wykaz_df = pd.read_excel("wykaz.xlsx", dtype=str)
    elementy_df = pd.read_excel("elementy1.xlsx", dtype=str)
    
    # Sprawdź czy są wymagane kolumny
    if 'Nazwa' not in wykaz_df.columns:
        print("❌ Błąd: wykaz.xlsx nie zawiera kolumny 'Nazwa'!")
        return False
    
    if 'Referencja' not in elementy_df.columns:
        print("❌ Błąd: elementy1.xlsx nie zawiera kolumny 'Referencja'!")
        return False
    
    # Przygotowanie kolumny wynikowej
    wykaz_df['propozycja'] = ""
    
    # Uzupełnij puste pola
    elementy_df['Referencja'] = elementy_df['Referencja'].fillna('')
    
    print(f"🔍 Rozpoczynam dopasowanie dla {len(wykaz_df)} pozycji...")
    dopasowane = 0
    
    # Iteracja po wykazie
    for idx, row in wykaz_df.iterrows():
        nazwa = row.get('Nazwa')
        if not isinstance(nazwa, str) or nazwa.strip() == "":
            continue
        
        nazwa = nazwa.strip()
        
        # Tworzymy regex wzorzec
        regex = re.escape(nazwa) + r'(?!\d)'
        
        # Szukamy dopasowań
        dopasowania = elementy_df[elementy_df['Referencja'].str.contains(regex, na=False, regex=True)]
        
        if not dopasowania.empty:
            match = dopasowania.iloc[0]
            wykaz_df.at[idx, 'propozycja'] = match.get('Referencja1', "")
            dopasowane += 1
    
    # Zapisz wynik
    wykaz_df.to_excel("propozycje.xlsx", index=False)
    print(f"✅ Dopasowanie zakończone!")
    print(f"   Dopasowano: {dopasowane}/{len(wykaz_df)} pozycji")
    print(f"   Wyniki zapisano w: propozycje.xlsx")
    return True


def main():
    """
    Główna funkcja - sprawdza czy wykaz.xlsx istnieje i decyduje co zrobić
    """
    print("=" * 50)
    print("  SKRYPT WYKAZ + PROPOZYCJE")
    print("=" * 50)
    
    if os.path.exists('wykaz.xlsx'):
        print("\n📋 Znaleziono plik wykaz.xlsx")
        print("   Uruchamiam dopasowanie propozycji...\n")
        uruchom_propozycje()
    else:
        print("\n📋 Nie znaleziono pliku wykaz.xlsx")
        print("   Tworzę nowy wykaz...\n")
        if utworz_wykaz():
            print("\n" + "=" * 50)
            print("ℹ️  NASTĘPNY KROK:")
            print("   1. Upewnij się, że masz plik 'elementy1.xlsx'")
            print("   2. Uruchom skrypt ponownie, aby dopasować propozycje")
            print("=" * 50)


if __name__ == "__main__":
    main()
