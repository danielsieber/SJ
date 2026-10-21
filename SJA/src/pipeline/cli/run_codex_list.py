import time
from pathlib import Path

from src.pipeline.phase_10_codex_list import get_scraper_files, export_file_list_to_excel


def main():
    start_time = time.time()

    # *****     start MODUL 10     *****     File Liste     ***************************
    dateien_liste = get_scraper_files()
    
    # Exportiere in Liste_Codex.xlsx im Repo-Root
    AlsExcelSpeichern = "J"
    if AlsExcelSpeichern == "J":  
        repo_root = Path(__file__).resolve().parent.parent.parent.parent
        output_path = repo_root / "Liste_Codex.xlsx"
        export_file_list_to_excel(dateien_liste, str(output_path))
    
    # Verarbeiten der zurückgegebenen Liste im Hauptprogramm
    AlsListeImTerminal = "N"
    if AlsListeImTerminal == "J":      
        print("\n_99-01 Folgende Dateien wurden gefunden und verarbeitet:")
        print()
        for datei in dateien_liste:
            print(f"- {datei['Filename']}")
    # *****     ENDE MODUL 10     ******************************************************

    end_time = time.time()
    dauer = end_time - start_time
    print(f"\nDas Programm dauerte {dauer:.3f} Sekunden.")


if __name__ == "__main__":
    main()
