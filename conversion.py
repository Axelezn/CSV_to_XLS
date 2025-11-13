import pandas as pd 
import os
import sys
import tkinter as tk
from tkinter import filedialog, simpledialog

# --- FENÊTRES DE DIALOGUE GRAPHIQUE ---
def afficher_dialogue_et_parametres():
    root = tk.Tk()
    root.withdraw() 
    #Dossier en entrée
    dossierSource = filedialog.askdirectory(
        title="Choisissez un dossier qui contient les CSV que vous voulez traiter"
    )
    
    if not dossierSource:
        print("Opération annulée. Aucun dossier sélectionné.")
        sys.exit()
        #Fichier en sortie
    nomExcel = simpledialog.askstring(
        "Nom du fichier Excel de sortie",
        "Entrez le nom du fichier Excel de sortie (ex: Rapport_Final.xlsx):",
        initialvalue="Rapport_Fusionne.xlsx"
    )
    if not nomExcel:
        print("Opération annulée. Aucun nom de fichier de sortie défini.")
        sys.exit()
    if not nomExcel.lower().endswith('.xlsx'):
        nomExcel += '.xlsx'
        
    return dossierSource, nomExcel

# --- RENOMMAGE AUTOMATIQUE ---
def renameCsv(dossierSource):
    nb_renommes = 0
    print("\n--- ÉTAPE DE PRÉPARATION : Nettoyage des noms de fichiers ---")
    
    for nomFichier in os.listdir(dossierSource):
        if nomFichier.endswith('.csv'):
            parties = nomFichier.replace('.csv', '').split('_')
            if len(parties) >= 2 and parties[-1].isdigit() and len(parties[-1]) == 4:
                # Nouveau nom : MOIS_ANNEE.csv
                nouveau_nom = f"{parties[-2]}_{parties[-1]}.csv"
                
                if nomFichier != nouveau_nom:
                    chemin_ancien = os.path.join(dossierSource, nomFichier)
                    chemin_nouveau = os.path.join(dossierSource, nouveau_nom)
                    
                    try:
                        os.rename(chemin_ancien, chemin_nouveau)
                        print(f"🔄 Renommé : {nomFichier} -> {nouveau_nom}")
                        nb_renommes += 1
                    except Exception as e:
                        print(f"❌ Erreur de renommage pour {nomFichier} : {e}")
            
    if nb_renommes == 0:
        print("Aucun fichier à renommer (les noms sont déjà courts ou le format n'est pas ID_MOIS_ANNEE).")


# --- FONCTION PRINCIPALE DE CONVERSION ET FUSION ---
def excelConverter(dossierSource, nomExcel, mode): 
    
    path_excel = os.path.join(dossierSource, nomExcel)
    
    ##Configs pour +eurs types de csv
    configurations = [(';', 'utf-8'), (';', 'latin-1'), (',', 'utf-8'), (',', 'latin-1')]
    dataframes = {}
    df_concat = pd.DataFrame()
    
    print("\n---------------------------------------------------------")
    print(f"Démarrage en mode : {'CONCATENATION (une seule feuille)' if mode == 'concat' else 'MULTI-PAGES (une feuille par fichier)'}")
    print("---------------------------------------------------------")

    for nomFichier in os.listdir(dossierSource) :
        if nomFichier.endswith('.csv') and os.path.isfile(os.path.join(dossierSource, nomFichier)): 
            chemin_csv = os.path.join(dossierSource, nomFichier)
            conversion_reussie = False
            
            for sep, encoding in configurations:
                try:
                    df = pd.read_csv(chemin_csv, sep=sep, encoding=encoding)
                    
                    # VÉRIFICATION CRITIQUE : S'assurer qu'il y a plus d'une colonne (séparateur correct)
                    if df.shape[1] > 1:
                        conversion_reussie = True
                        
                        if mode == 'concat':
                            # Concaténation
                            df['Fichier_Source'] = nomFichier.replace('.csv', '')
                            df_concat = pd.concat([df_concat, df], ignore_index=True)
                            print(f"✅ Lecture réussie de {nomFichier} (Ajouté à la feuille unique)")
                        else:
                            # Multi-Pages
                            nom_feuille = nomFichier.replace('.csv', '') 
                            dataframes[nom_feuille] = df
                            print(f"✅ Lecture réussie de {nomFichier} (Feuille: {nom_feuille})")
                        
                        break 
                except Exception:
                    continue 
            
            if not conversion_reussie:
                print(f"⚠️ Échec de la lecture pour {nomFichier}.")
    
    # --- ÉTAPE D'ÉCRITURE FINALE ---
    
    if (mode == 'multi' and dataframes) or (mode == 'concat' and not df_concat.empty):
        print("\n---------------------------------------------------------")
        print(f"Écriture du fichier Excel : {nomExcel}...")
        try:
            with pd.ExcelWriter(path_excel, engine='xlsxwriter') as writer:
                if mode == 'concat':
                    df_concat.to_excel(writer, sheet_name='Fusion_Totale', index=False)
                else:
                    for sheet_name, df in dataframes.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"🎉 Succès ! Le fichier Excel a été généré dans : {os.path.dirname(path_excel) or dossierSource}")
            
        except Exception as e:
            print(f"❌ Erreur lors de l'écriture du fichier Excel : {e}")
    else:
        print("Aucun fichier CSV valide trouvé ou traité.")


# --- DÉMARRAGE DU PROGRAMME ---
if __name__ == "__main__":
    
    # Instructions Console
    print("=======================================================")
    print("         CONVERTISSEUR CSV VERS EXCEL")
    print("=======================================================")
    print("ATTENTION : Le fichier Excel de sortie sera créé dans le dossier source.")
    print("Veuillez suivre les étapes dans les fenêtres de dialogue qui vont apparaître.")
    print("-------------------------------------------------------")
    
    
    dossierSource, nomExcel = afficher_dialogue_et_parametres()
    
    
    renameCsv(dossierSource)
    
    
    print("\n=======================================================")
    print("Choisissez le mode de conversion :")
    print("1 - Fichiers séparés : 1 CSV = 1 feuille Excel")
    print("2 - Concaténation : Tous les CSV sur 1 seule feuille")
    
    while True:
        choix = input("Entrez 1 ou 2 : ").strip()
        if choix == '1':
            mode_execution = 'multi'
            break
        elif choix == '2':
            mode_execution = 'concat'
            break
        else:
            print("Choix invalide. Veuillez entrer '1' ou '2'.")
            
    
    excelConverter(dossierSource, nomExcel, mode_execution)
    
    # Pause finale
    input("\nConversion terminée. Appuyez sur Entrée pour quitter...")
