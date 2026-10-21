import sys
import time
import os
from pathlib import Path

from src.pipeline.phase_20_preprocessing import (
    main_analysis,
    repair_data,
    clean_msg_data,
    clean_fol_data,
    format_and_save,
)


def main():
    # Clear terminal
    os.system('cls' if os.name == 'nt' else 'clear')
    
    # Start
    print("\n_20-01 Starte Raw-Daten Analyse...")
    
    # 1. Eingabe des Weekends
    if len(sys.argv) > 1:
        weekend_input = sys.argv[1]
        print(f"\nWeekend aus Argument geladen: {weekend_input}")
    else:
        try:
            weekend_input = input("\nWelches Weekend möchtest du auswerten? ")
        except EOFError:
            print("\nFehler: Kein interaktives Terminal gefunden und kein Argument übergeben.")
            print("Verwendung: python -m src.pipeline.cli.run_analysis <Weekend>")
            return

    try:
        weekend_id = int(weekend_input)
    except ValueError:
        weekend_id = weekend_input

    # Start Zeitmessung nach Input
    start_time = time.time()
    
    repo_root = Path(__file__).resolve().parent.parent.parent.parent
    root_dir = repo_root
    print("Root Directory root_dir:  ", root_dir)

    # *****     start MODUL 20     *****     import raw data    ************************
    result = main_analysis(weekend_id, str(root_dir))
    # *****     end MODUL 20       *****************************************************

    if result:
        fol_rawdata, msg_rawdata, output_filename, output_path = result
        
        # *****     start MODUL 22     *****     repair data        ************************
        print("\n_22-00 Starte Reparatur der fol_rawdata...")
        fol_repairdata = repair_data(fol_rawdata)
        # *****     end MODUL 22       ******************************************************
        
        # *****     start MODUL 24     *****     clean data         ************************
        print("\n_24-00 Starte Bereinigung der msg_rawdata und fol_repairdata...")
        msg_cleandata = clean_msg_data(msg_rawdata)
        fol_cleandata = clean_fol_data(fol_repairdata)
        # *****     end MODUL 24       ******************************************************
        
        # *****     start MODUL 28     *****     format and save    ************************
        print("\n_28-00 Starte Formatierung und Speicherung...")
        format_and_save(
            fol_rawdata,
            msg_rawdata,
            msg_cleandata,
            fol_repairdata,
            fol_cleandata,
            output_filename,
            output_path,
            pdf_output_path=root_dir / "Data_pdf",
        )
        # *****     end MODUL 28     ******************************************************

    # *****     ENDE  *****  

    end_time = time.time()
    dauer = end_time - start_time
    print(f"\nDas Programm dauerte {dauer:.3f} Sekunden.")

if __name__ == "__main__":
    main()
