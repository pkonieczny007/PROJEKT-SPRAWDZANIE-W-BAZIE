import subprocess
import shutil
import os

def main():
    print("Instalacja PROJEKT SPRAWDZENIE W BAZIE...")

    # Klonowanie repozytorium
    repo_url = "https://github.com/pkonieczny007/PROJEKT-SPRAWDZANIE-W-BAZIE.git"
    clone_dir = "PROJEKT-SPRAWDZANIE-W-BAZIE"

    print(f"Klonowanie repozytorium z {repo_url}...")
    result = subprocess.run(["git", "clone", repo_url, clone_dir], capture_output=True, text=True)

    if result.returncode == 0:
        print("✅ Repozytorium zostało pomyślnie sklonowane.")
    else:
        print("❌ Błąd podczas klonowania repozytorium:")
        print(result.stderr)
        return

    # Przeniesienie pliku setup.py do folderu TMP
    setup_path = "setup/setup.py"
    tmp_path = "TMP/setup.py"

    if os.path.exists(setup_path):
        print("Przenoszenie pliku instalacyjnego do TMP...")
        shutil.move(setup_path, tmp_path)
        print("✅ Plik setup.py przeniesiony do TMP.")
    else:
        print("⚠️  Plik setup.py nie znaleziony w setup/.")

    print("Instalacja zakończona.")

if __name__ == "__main__":
    main()
