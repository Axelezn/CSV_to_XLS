import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

class ApplicationConvertisseur:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur de fichiers CSV vers Excel - DEMB")
        self.root.geometry("700x650")

        # --- VARIABLES DE STOCKAGE ---
        self.dossier_source_path = tk.StringVar()
        # CORRECTION 1 : "fhichier" corrigé en "fichier" pour correspondre au reste du code
        self.nom_fichier_sortie = tk.StringVar(value="Rapport_Fusionne.xlsx")
        self.mode_execution = tk.StringVar(value="multi") 
        
        # --- Interfaçage ---
        self.creer_widgets()

    def creer_widgets(self): 
        # 1ER BLOC : SELECTION DOSSIER
        # CORRECTION 2 : "pasx" corrigé en "padx"
        frame_dossier = tk.LabelFrame(self.root, text="Etape 1 : Sélectionnez le dossier où sont les CSV", padx=10, pady=10)
        frame_dossier.pack(fill="x", padx=10, pady=5)

        lbl_dossier = tk.Label(frame_dossier, text="Dossier contenant les fichiers CSV :")
        lbl_dossier.pack(anchor="w")

        frame_input_dossier = tk.Frame(frame_dossier)
        frame_input_dossier.pack(fill="x")

        entry_dossier = tk.Entry(frame_input_dossier, textvariable=self.dossier_source_path, width=50)
        entry_dossier.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # CORRECTION 3 : Le bouton était mal configuré (mauvais texte et manque la commande)
        btn_browse = tk.Button(frame_input_dossier, text="Parcourir...", command=self.choisir_dossier)
        btn_browse.pack(side="right")

        # 2EME BLOC : PARAMETRES DE SORTIE
        frame_params = tk.LabelFrame(self.root, text="2. Paramètre de conversion", padx=10, pady=10)
        frame_params.pack(fill="x", padx=10, pady=5)

        # NOM FICHIER DE SORTIE
        lbl_nom = tk.Label(frame_params, text="Nom du fichier Excel de sortie : ")
        lbl_nom.grid(row=0, column=0, sticky="w", pady=5)
        
        # CORRECTION 4 : Utilisation du bon nom de variable corrigé plus haut
        entry_nom = tk.Entry(frame_params, textvariable=self.nom_fichier_sortie, width=40)
        entry_nom.grid(row=0, column=1, sticky="w", padx=10)

        # MODE (Radio Buttons)
        lbl_mode = tk.Label(frame_params, text="Mode de conversion :")
        lbl_mode.grid(row=1, column=0, sticky="w", pady=10)

        frame_radios = tk.Frame(frame_params)
        frame_radios.grid(row=1, column=1, sticky="w", padx=10)

        # Création des boutons radio avec liaison variable
        rb1 = tk.Radiobutton(frame_radios, text="Mode Multi-Feuilles (1 CSV = 1 Feuille Excel)", 
                             variable=self.mode_execution, value="multi")
        rb1.pack(anchor="w")

        rb2 = tk.Radiobutton(frame_radios, text="Mode Addition : Tous les CSV sur une seule feuille Excel", 
                             variable=self.mode_execution, value="concat")
        rb2.pack(anchor="w")

        # 3. Boutons d'actions
        self.btn_action = tk.Button(self.root, text="LANCER LA CONVERSION", 
                                    bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
                                    command=self.lancer_processus, height=2)
        self.btn_action.pack(fill="x", padx=10, pady=20)

        # ZONE DE LOGS 
        lbl_log = tk.Label(self.root, text="Journal d'exécution :")
        lbl_log.pack(anchor="w", padx=10)

        self.log_area = scrolledtext.ScrolledText(self.root, state='disabled', font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # --- FONCTIONS UTILITAIRES GUI --- 
    def choisir_dossier(self):
        dossier = filedialog.askdirectory()
        if dossier:
            self.dossier_source_path.set(dossier)
            self.log(f"Dossier sélectionné : {dossier}")

    def log(self, message, tag=None):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n", tag)
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update() 

    def lancer_processus(self):
        dossier = self.dossier_source_path.get()
        nom_sortie = self.nom_fichier_sortie.get()
        mode = self.mode_execution.get()

        # Validations de base
        if not dossier or not os.path.exists(dossier):
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier valide.")
            return
        
        if not nom_sortie:
            messagebox.showerror("Erreur", "Veuillez donner un nom au fichier Excel.")
            return

        if not nom_sortie.lower().endswith('.xlsx'):
            nom_sortie += '.xlsx'
            self.nom_fichier_sortie.set(nom_sortie)

        # Désactiver le bouton pour éviter le double-clic
        self.btn_action.config(state="disabled", text="Traitement en cours...")
        self.log("-" * 50)
        self.log("DÉMARRAGE DU TRAITEMENT")
        
        try:
            # 1. Renommage
            self.renommer_fichiers_csv(dossier)
            
            # 2. Conversion
            self.convertir_csv_en_excel_fusion(dossier, nom_sortie, mode)
            
            messagebox.showinfo("Succès", "Traitement terminé avec succès !")
            
        except Exception as e:
            self.log(f"ERREUR CRITIQUE : {e}")
            messagebox.showerror("Erreur", f"Une erreur est survenue :\n{e}")
        
        finally:
            self.btn_action.config(state="normal", text="LANCER LA CONVERSION")

    def renommer_fichiers_csv(self, dossier_source):
        self.log("--- Étape 1 : Vérification des noms de fichiers ---")
        nb_renommes = 0
        try:
            for nom_fichier in os.listdir(dossier_source):
                if nom_fichier.endswith('.csv'):
                    parties = nom_fichier.replace('.csv', '').split('_')
                    if len(parties) >= 2 and parties[-1].isdigit() and len(parties[-1]) == 4:
                        nouveau_nom = f"{parties[-2]}_{parties[-1]}.csv"
                        if nom_fichier != nouveau_nom:
                            chemin_ancien = os.path.join(dossier_source, nom_fichier)
                            chemin_nouveau = os.path.join(dossier_source, nouveau_nom)
                            try:
                                os.rename(chemin_ancien, chemin_nouveau)
                                self.log(f"🔄 Renommé : {nom_fichier} -> {nouveau_nom}")
                                nb_renommes += 1
                            except Exception as e:
                                self.log(f"❌ Erreur renommage {nom_fichier} : {e}")
        except Exception as e:
             self.log(f"Erreur lecture dossier : {e}")

        if nb_renommes == 0:
            self.log("Aucun fichier nécessitant un renommage trouvé.")

    def convertir_csv_en_excel_fusion(self, dossier_source, nom_fichier_excel, mode):
        path_excel = os.path.join(dossier_source, nom_fichier_excel)
        configurations = [(';', 'utf-8'), (';', 'latin-1'), (',', 'utf-8'), (',', 'latin-1')]
        dataframes = {}
        df_concat = pd.DataFrame()

        self.log("\n--- Étape 2 : Lecture et Conversion ---")
        
        fichiers_csv = [f for f in os.listdir(dossier_source) if f.endswith('.csv')]
        if not fichiers_csv:
            self.log("⚠️ Aucun fichier CSV trouvé dans le dossier !")
            return

        for nom_fichier in fichiers_csv:
            chemin_csv = os.path.join(dossier_source, nom_fichier)
            conversion_reussie = False
            
            for sep, encoding in configurations:
                try:
                    df = pd.read_csv(chemin_csv, sep=sep, encoding=encoding)
                    if df.shape[1] > 1:
                        conversion_reussie = True
                        if mode == 'concat':
                            df['Fichier_Source'] = nom_fichier.replace('.csv', '')
                            df_concat = pd.concat([df_concat, df], ignore_index=True)
                            self.log(f"✅ {nom_fichier} : Ajouté (Concat)")
                        else:
                            nom_feuille = nom_fichier.replace('.csv', '')[:30] 
                            dataframes[nom_feuille] = df
                            self.log(f"✅ {nom_fichier} : Lu (Feuille: {nom_feuille})")
                        break 
                except Exception:
                    continue
            
            if not conversion_reussie:
                self.log(f"⚠️ Échec lecture : {nom_fichier}")

        # Ecriture
        if (mode == 'multi' and dataframes) or (mode == 'concat' and not df_concat.empty):
            self.log(f"\n--- Étape 3 : Écriture du fichier Excel ---")
            self.log(f"Création de {nom_fichier_excel}...")
            try:
                with pd.ExcelWriter(path_excel, engine='xlsxwriter') as writer:
                    if mode == 'concat':
                        df_concat.to_excel(writer, sheet_name='Fusion_Totale', index=False)
                    else:
                        for sheet_name, df in dataframes.items():
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                self.log(f"🎉 TERMINÉ ! Fichier créé dans le dossier source.")
            except Exception as e:
                self.log(f"❌ Erreur écriture Excel : {e}")
        else:
            self.log("❌ Aucune donnée valide à écrire.")

# --- LANCEMENT ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ApplicationConvertisseur(root)
    root.mainloop()
