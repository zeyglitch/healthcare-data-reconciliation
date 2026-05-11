import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess
import sys
from pathlib import Path

class AppConciliation:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle d'exhaustivite - DIM")
        self.root.geometry("1200x800")
        self.root.resizable(False, False)
        
        # Couleurs
        bg_color = "#f0f4f8"
        header_color = "#2F5496"
        btn_color = "#2F5496"
        btn_hover = "#1a3a6e"
        
        self.root.configure(bg=bg_color)
        
        # Variables
        self.fichier_orbis = tk.StringVar()
        self.fichiers_hexa = {}
        self.dossier_export = tk.StringVar(value=str(Path('./data_test/export_test').resolve()))
        
        noms_hexa = [
            ("Hexa Hospitalisations", "hexa_hospit"),
            ("Hexa Nouveau-Ne", "hexa_nn"),
            ("Hexa Orthogenie", "hexa_ortho"),
            ("Hexa Chimio", "hexa_chimio"),
            ("Hexa Dialyse P", "hexa_dialyse"),
            ("Hexa Hemodialyse", "hexa_hemo"),
        ]
        for _, key in noms_hexa:
            self.fichiers_hexa[key] = tk.StringVar()
        
        # ==========================================
        # EN-TÊTE
        # ==========================================
        header_frame = tk.Frame(root, bg=header_color, height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(
            header_frame, text="Controle d'exhaustivite - Conciliation DIM",
            font=("Segoe UI", 16, "bold"), fg="white", bg=header_color
        ).pack(pady=15)
        
        # ==========================================
        # CORPS PRINCIPAL
        # ==========================================
        main_frame = tk.Frame(root, bg=bg_color, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Section Orbis
        self._creer_section_fichier(main_frame, "Fichier Orbis :", self.fichier_orbis, 0)
        
        # Séparateur
        ttk.Separator(main_frame, orient="horizontal").grid(row=1, column=0, columnspan=3, sticky="ew", pady=8)
        
        # Section Hexa
        row_idx = 2
        for label, key in noms_hexa:
            self._creer_section_fichier(main_frame, f"{label} :", self.fichiers_hexa[key], row_idx)
            row_idx += 1
        
        # Séparateur
        ttk.Separator(main_frame, orient="horizontal").grid(row=row_idx, column=0, columnspan=3, sticky="ew", pady=8)
        row_idx += 1
        
        # Dossier export
        tk.Label(main_frame, text="Dossier export :", font=("Segoe UI", 9, "bold"), bg=bg_color, anchor="w").grid(row=row_idx, column=0, sticky="w", pady=2)
        tk.Entry(main_frame, textvariable=self.dossier_export, width=55, font=("Segoe UI", 8), state="readonly").grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)
        tk.Button(main_frame, text="...", width=3, command=lambda: self._choisir_dossier(self.dossier_export)).grid(row=row_idx, column=2, pady=2)
        row_idx += 1
        
        main_frame.columnconfigure(1, weight=1)
        
        # ==========================================
        # BARRE DE PROGRESSION + BOUTON
        # ==========================================
        bottom_frame = tk.Frame(root, bg=bg_color, padx=20, pady=10)
        bottom_frame.pack(fill=tk.X)
        
        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate', length=400)
        self.progress.pack(side=tk.LEFT, padx=(0, 15))
        
        self.status_label = tk.Label(bottom_frame, text="Pret.", font=("Segoe UI", 9), bg=bg_color, fg="#555")
        self.status_label.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        self.btn_lancer = tk.Button(
            bottom_frame, text="Lancer le traitement", font=("Segoe UI", 11, "bold"),
            bg=btn_color, fg="white", activebackground=btn_hover, activeforeground="white",
            relief="flat", padx=15, pady=5, cursor="hand2",
            command=self._lancer_traitement
        )
        self.btn_lancer.pack(side=tk.RIGHT)
    
    def _creer_section_fichier(self, parent, label, var, row):
        bg_color = "#f0f4f8"
        tk.Label(parent, text=label, font=("Segoe UI", 9, "bold"), bg=bg_color, anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        tk.Entry(parent, textvariable=var, width=55, font=("Segoe UI", 8), state="readonly").grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        tk.Button(parent, text="...", width=3, command=lambda: self._choisir_fichier(var)).grid(row=row, column=2, pady=2)
    
    def _choisir_fichier(self, var):
        fichier = filedialog.askopenfilename(
            title="Selectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if fichier:
            var.set(fichier)
    
    def _choisir_dossier(self, var):
        dossier = filedialog.askdirectory(title="Selectionner le dossier d'export")
        if dossier:
            var.set(dossier)

    def _valider_groupe_hexagone(self, valeurs, libelle_groupe):
        fichiers_selectionnes = [valeur for valeur in valeurs if valeur]
        if not fichiers_selectionnes:
            return True, ""
        if len(fichiers_selectionnes) != len(valeurs):
            return False, f"Le groupe '{libelle_groupe}' doit contenir les 3 fichiers ou être vide."
        return True, ""
    
    def _lancer_traitement(self):
        # Vérifications
        if not self.fichier_orbis.get():
            messagebox.showwarning("Attention", "Veuillez selectionner le fichier Orbis.")
            return

        hospit_ok, msg_hospit = self._valider_groupe_hexagone(
            [self.fichiers_hexa["hexa_hospit"].get(), self.fichiers_hexa["hexa_nn"].get(), self.fichiers_hexa["hexa_ortho"].get()],
            "Hexa Hospitalisations"
        )
        if not hospit_ok:
            messagebox.showwarning("Attention", msg_hospit)
            return

        seances_ok, msg_seances = self._valider_groupe_hexagone(
            [self.fichiers_hexa["hexa_chimio"].get(), self.fichiers_hexa["hexa_dialyse"].get(), self.fichiers_hexa["hexa_hemo"].get()],
            "Hexa Séances"
        )
        if not seances_ok:
            messagebox.showwarning("Attention", msg_seances)
            return

        if not any([
            self.fichiers_hexa["hexa_hospit"].get(), self.fichiers_hexa["hexa_nn"].get(), self.fichiers_hexa["hexa_ortho"].get(),
            self.fichiers_hexa["hexa_chimio"].get(), self.fichiers_hexa["hexa_dialyse"].get(), self.fichiers_hexa["hexa_hemo"].get()
        ]):
            messagebox.showwarning("Attention", "Veuillez selectionner les 3 fichiers Hexagone d'un groupe complet ou les 2 groupes complets.")
            return
        
        # Préparer les chemins pour le script
        self.btn_lancer.config(state="disabled")
        self.progress.start(15)
        self.status_label.config(text="Traitement en cours...")
        
        # Lancer dans un thread séparé pour ne pas bloquer l'interface
        thread = threading.Thread(target=self._executer_script, daemon=True)
        thread.start()
    
    def _executer_script(self):
        try:
            # Importer le module de conciliation
            import controle_exhaust
            
            # Récupérer les chemins
            orbis_path = self.fichier_orbis.get()
            hexa_hospit = [p for p in [self.fichiers_hexa["hexa_hospit"].get(), 
                                     self.fichiers_hexa["hexa_nn"].get(), 
                                     self.fichiers_hexa["hexa_ortho"].get()] if p]
            hexa_seances = [p for p in [self.fichiers_hexa["hexa_chimio"].get(), 
                                      self.fichiers_hexa["hexa_dialyse"].get(), 
                                      self.fichiers_hexa["hexa_hemo"].get()] if p]
            export_dir = self.dossier_export.get()
            
            # Lancer directement la fonction Python
            controle_exhaust.lancer_conciliation(
                orbis_path=orbis_path,
                hexa_hospit_paths=hexa_hospit,
                hexa_seances_paths=hexa_seances,
                export_dir=export_dir
            )
            
            # Si aucune exception n'est levée, c'est un succès
            self.root.after(0, self._traitement_termine, True, "")
        except Exception as e:
            # En cas d'erreur, on capture l'exception
            import traceback
            error_msg = traceback.format_exc()
            self.root.after(0, self._traitement_termine, False, error_msg)
    
    def _traitement_termine(self, success, error_msg):
        self.progress.stop()
        self.btn_lancer.config(state="normal")
        
        if success:
            self.status_label.config(text="Traitement termine avec succes !", fg="#2e7d32")
            messagebox.showinfo("Succes", f"Traitement termine !\n\nFichiers generes dans :\n{self.dossier_export.get()}")
        else:
            self.status_label.config(text="Erreur lors du traitement.", fg="#c62828")
            # Afficher seulement la fin du message d'erreur si trop long
            short_error = error_msg[-500:] if len(error_msg) > 500 else error_msg
            messagebox.showerror("Erreur", f"Le traitement a echoue :\n\n{short_error}")
    
    def _traitement_erreur(self, msg):
        self.progress.stop()
        self.btn_lancer.config(state="normal")
        self.status_label.config(text="Erreur.", fg="#c62828")
        messagebox.showerror("Erreur", f"Erreur inattendue :\n{msg}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppConciliation(root)
    root.mainloop()
