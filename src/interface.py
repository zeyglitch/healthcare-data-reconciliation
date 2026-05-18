"""
Interface graphique gérant les inputs utilisateur pour le script de conciliation.
Construite avec Tkinter (standard Python) et utilise un Threading basique pour ne pas freezer l'UI
lors du traitement de grosses bases de données.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path
import controle_exhaust

class AppConciliation:
    """
    Classe englobant toute l'interface. On utilise une classe pour partager 
    l'état (les variables de chemins de fichiers, les labels d'état) entre toutes les fonctions.
    """
    def __init__(self, root):
        self.root = root  # Fenêtre principale (racine de l'arbre Tkinter)
        self.root.title("Controle d'exhaustivite - DIM")
        
        # Dimensions par défaut et minimales pour éviter un layout cassé
        self.root.geometry("1220x820")
        self.root.minsize(980, 680)
        self.root.resizable(True, True)
        
        # Couleurs
        bg_color = "#eef3f9"
        panel_color = "#ffffff"
        header_color = "#2F5496"
        btn_color = "#2F5496"
        btn_hover = "#1a3a6e"
        muted_text = "#5b6470"
        
        self.root.configure(bg=bg_color)

        # --- Personnalisation du style (Thème) ---
        # On utilise ttk (Themed Tkinter) pour avoir des composants un peu plus modernes.
        style = ttk.Style(self.root)
        try:
            # "clam" est un thème intégré à Tkinter qui fait moins "Windows 95" que le style par défaut.
            style.theme_use("clam")
        except tk.TclError:
            pass
            
        # Configuration des couleurs pour nos différents composants (cadres, textes, etc.)
        style.configure("App.TFrame", background=bg_color)
        style.configure("Panel.TFrame", background=panel_color)
        style.configure("App.TLabel", background=bg_color, foreground="#17212b")
        style.configure("Muted.TLabel", background=bg_color, foreground=muted_text)
        style.configure("Panel.TLabel", background=panel_color, foreground="#17212b")
        style.configure("Section.TLabelframe", background=panel_color, borderwidth=1, relief="solid")
        style.configure("Section.TLabelframe.Label", background=panel_color, foreground="#17212b", font=("Segoe UI", 10, "bold"))
        style.configure("App.Horizontal.TProgressbar", troughcolor="#d9e2ef", background=btn_color, bordercolor="#d9e2ef", lightcolor=btn_color, darkcolor=btn_color)

        self.default_entry_width = 72
        self.browse_width = 4
        
        # --- Variables d'état Tkinter (StringVar) ---
        # Les StringVar sont spéciales : si elles changent dans le code, l'interface se met à jour
        # automatiquement, et si l'utilisateur saisit du texte, la variable se met à jour.
        self.fichier_orbis = tk.StringVar()
        self.fichiers_hexa = {}
        
        # On définit un dossier d'export par défaut. 
        # Path(...).resolve() trouve le chemin absolu propre sur l'ordinateur, peu importe l'OS.
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
        header_frame = tk.Frame(root, bg=header_color, height=72)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(
            header_frame, text="Controle d'exhaustivite - Conciliation DIM",
            font=("Segoe UI Semibold", 18), fg="white", bg=header_color
        ).pack(pady=18)
        
        # ==========================================
        # CORPS PRINCIPAL
        # ==========================================
        content_host = tk.Frame(root, bg=bg_color)
        content_host.pack(fill=tk.BOTH, expand=True)

        # --- Système de défilement (Scrollbar) ---
        # Tkinter ne sait pas faire défiler une fenêtre (Frame) classique nativement.
        # L'astuce classique est de créer un "Canvas" (une zone de dessin qui, elle, peut défiler), 
        # d'y attacher une barre de défilement, et de dessiner notre Frame à l'intérieur de ce Canvas.
        self.canvas = tk.Canvas(content_host, bg=bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_host, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # .pack() est un gestionnaire de layout simple (empilement)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.main_frame = tk.Frame(self.canvas, bg=bg_color, padx=20, pady=18)
        self.main_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._ajuster_largeur_canvas)
        self.root.bind("<MouseWheel>", self._defiler_souris)

        info_bar = tk.Frame(self.main_frame, bg=bg_color)
        info_bar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        tk.Label(
            info_bar,
            text="Sélectionnez un fichier Orbis, puis un groupe Hexagone complet ou les deux groupes complets.",
            font=("Segoe UI", 9),
            bg=bg_color,
            fg=muted_text,
            anchor="w",
            justify="left"
        ).pack(fill=tk.X)

        # --- Panneau principal des fichiers ---
        files_panel = tk.Frame(self.main_frame, bg=panel_color, padx=16, pady=16, bd=1, relief="solid")
        
        # .grid() place les éléments dans une grille (ligne/colonne). 
        # sticky="nsew" dit au composant de s'étirer (Nord, Sud, Est, Ouest) pour prendre toute la place.
        files_panel.grid(row=1, column=0, sticky="nsew")
        
        # weight=1 permet à la colonne/ligne de s'agrandir si on redimensionne la fenêtre
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)

        # tk.LabelFrame crée un cadre esthétique avec un titre intégré en haut à gauche. Très propre pour grouper.
        orbis_section = tk.LabelFrame(files_panel, text="Fichier Orbis", bg=panel_color, fg="#17212b", padx=12, pady=12)
        orbis_section.grid(row=0, column=0, sticky="ew") # "ew" (Est-West) = s'étire seulement horizontalement
        orbis_section.columnconfigure(1, weight=1)
        self._creer_section_fichier(orbis_section, "Orbis", self.fichier_orbis, 0)

        hospital_section = tk.LabelFrame(files_panel, text="Hexagone - Hospitalisations", bg=panel_color, fg="#17212b", padx=12, pady=12)
        hospital_section.grid(row=1, column=0, sticky="ew", pady=(14, 10))
        hospital_section.columnconfigure(1, weight=1)

        row_idx = 0
        for label, key in noms_hexa[:3]:
            self._creer_section_fichier(hospital_section, label, self.fichiers_hexa[key], row_idx)
            row_idx += 1

        seance_section = tk.LabelFrame(files_panel, text="Hexagone - Séances", bg=panel_color, fg="#17212b", padx=12, pady=12)
        seance_section.grid(row=2, column=0, sticky="ew", pady=(6, 10))
        seance_section.columnconfigure(1, weight=1)

        row_idx = 0
        for label, key in noms_hexa[3:]:
            self._creer_section_fichier(seance_section, label, self.fichiers_hexa[key], row_idx)
            row_idx += 1

        export_section = tk.LabelFrame(files_panel, text="Export", bg=panel_color, fg="#17212b", padx=12, pady=12)
        export_section.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        export_section.columnconfigure(1, weight=1)
        self._creer_section_export(export_section, 0)

        files_panel.columnconfigure(0, weight=1)
        
        # ==========================================
        # BARRE DE PROGRESSION + BOUTON
        # ==========================================
        bottom_frame = tk.Frame(root, bg=bg_color, padx=20, pady=12)
        bottom_frame.pack(fill=tk.X)
        
        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate', length=360, style="App.Horizontal.TProgressbar")
        self.progress.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True)
        
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
        """
        Fonction utilitaire pour éviter de répéter 10 fois le même code de création de bouton/champ texte.
        Génère une ligne (row) avec : un Titre (Label), un champ texte (Entry) et un bouton Parcourir (Button).
        """
        bg_color = parent.cget("bg")
        tk.Label(parent, text=label, font=("Segoe UI", 9, "bold"), bg=bg_color, anchor="w").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
        
        # L'Entry est en "readonly" pour empêcher l'utilisateur de taper n'importe quoi au clavier. Il DOIT utiliser le bouton.
        tk.Entry(parent, textvariable=var, width=self.default_entry_width, font=("Segoe UI", 9), state="readonly").grid(row=row, column=1, sticky="ew", padx=5, pady=6)
        
        # Le paramètre 'command=lambda: ...' permet d'associer le clic du bouton à notre fonction de choix de fichier.
        tk.Button(parent, text="...", width=self.browse_width, command=lambda: self._choisir_fichier(var)).grid(row=row, column=2, pady=6, padx=(6, 0))

    def _creer_section_export(self, parent, row):
        bg_color = parent.cget("bg")
        tk.Label(parent, text="Dossier export", font=("Segoe UI", 9, "bold"), bg=bg_color, anchor="w").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
        tk.Entry(parent, textvariable=self.dossier_export, width=self.default_entry_width, font=("Segoe UI", 9), state="readonly").grid(row=row, column=1, sticky="ew", padx=5, pady=6)
        tk.Button(parent, text="...", width=self.browse_width, command=lambda: self._choisir_dossier(self.dossier_export)).grid(row=row, column=2, pady=6, padx=(6, 0))
    
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

    def _chemins_uniques(self, chemins):
        chemins_uniques = []
        vus = set()
        for chemin in chemins:
            if chemin and chemin not in vus:
                chemins_uniques.append(chemin)
                vus.add(chemin)
        return chemins_uniques

    def _ajuster_largeur_canvas(self, event):
        self.canvas.itemconfigure(self.main_window, width=event.width)

    def _defiler_souris(self, event):
        if self.canvas.winfo_height() < self.main_frame.winfo_reqheight():
            self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def _valider_groupe_hexagone(self, valeurs, libelle_groupe):
        """
        Règle de validation métier pour l'IHM : 
        On s'assure qu'un groupe de fichiers (Hospit ou Séances) est soit complètement vide (0 fichier),
        soit complètement rempli (3 fichiers différents). S'il en manque un, on bloque.
        """
        # On crée une liste en filtrant les chemins vides
        fichiers_selectionnes = [valeur for valeur in valeurs if valeur]
        
        if not fichiers_selectionnes:
            # Le groupe est totalement vide. C'est autorisé (on analysera juste l'autre groupe).
            return True, ""
            
        # Si on convertit une liste en "set" (ensemble mathématique), les doublons sont détruits.
        # Donc si la taille de la liste est différente du set, c'est que l'utilisateur a mis 2 fois le même fichier.
        if len(set(fichiers_selectionnes)) != len(fichiers_selectionnes):
            return False, f"Le groupe '{libelle_groupe}' contient le même fichier plusieurs fois. Choisissez des fichiers différents."
            
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
        self.root.update_idletasks()
        
        # --- THREADING (IMPORTANT) ---
        # Si on lance l'analyse Pandas directement ici, l'interface Tkinter va bloquer (freeze)
        # en attendant la fin du calcul, car on est sur le Main Thread.
        # On délègue donc l'exécution à un thread secondaire.
        # daemon=True permet au thread de s'arrêter proprement si on ferme la fenêtre.
        thread = threading.Thread(target=self._executer_script, daemon=True)
        thread.start()
    
    def _executer_script(self):
        """
        Méthode exécutée en arrière-plan. Elle fait l'interface avec notre logique métier.
        """
        try:
            # On importe notre script d'analyse métier à la volée.
            import controle_exhaust
            
            # Récupération des chemins actuels stockés dans nos variables Tkinter
            orbis_path = self.fichier_orbis.get()
            hexa_hospit = self._chemins_uniques([
                self.fichiers_hexa["hexa_hospit"].get(),
                self.fichiers_hexa["hexa_nn"].get(),
                self.fichiers_hexa["hexa_ortho"].get(),
            ])
            hexa_seances = self._chemins_uniques([
                self.fichiers_hexa["hexa_chimio"].get(),
                self.fichiers_hexa["hexa_dialyse"].get(),
                self.fichiers_hexa["hexa_hemo"].get(),
            ])
            export_dir = self.dossier_export.get()
            
            # Lancer directement la fonction Python
            controle_exhaust.lancer_conciliation(
                orbis_path=orbis_path,
                hexa_hospit_paths=hexa_hospit,
                hexa_seances_paths=hexa_seances,
                export_dir=export_dir
            )
            
            # Tout s'est bien passé.
            # ATTENTION : on ne modifie jamais l'interface Tkinter depuis un thread secondaire.
            # root.after(0, func, args) met func dans la file d'attente du thread principal 
            # pour qu'il l'exécute le plus vite possible (délai 0).
            self.root.after(0, self._traitement_termine, True, "")
        except Exception as e:
            # Si l'analyse plante, on catch l'erreur pour ne pas crasher l'application.
            import traceback
            error_msg = traceback.format_exc()  # Traceback donne la pile d'appels, essentiel pour débugger.
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


if __name__ == "__main__":
    root = tk.Tk()
    app = AppConciliation(root)
    root.mainloop()
