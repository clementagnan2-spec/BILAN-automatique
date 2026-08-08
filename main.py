# -*- coding: utf-8 -*-
"""
main.py — États Financiers Automatiques
Interface graphique : importer Balance N-1 + Balance N, générer le Bilan,
le Compte de Résultat, la Situation Financière et le Flux de Trésorerie à
partir de leurs modèles Excel respectifs, en appliquant les formules
CtaCptSolde... et les rubriques [xxx.EtLoc].
"""

import os
import sys
import traceback
import webbrowser

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import core

APP_TITLE = "États Financiers Automatiques"
APP_VERSION = "2.0"


def resource_path(relative_path: str) -> str:
    """Résout le chemin d'une ressource, y compris depuis un .exe PyInstaller
    (onefile) où les fichiers sont extraits dans un dossier temporaire."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class EtatsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} — v{APP_VERSION}")
        self.geometry("780x640")
        self.minsize(700, 560)

        self.balance_n1_path = tk.StringVar()
        self.balance_n_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.status_var = tk.StringVar(value="Prêt.")

        # une case à cocher + un champ "modèle personnalisé" par état
        self.etat_vars = {}       # id -> BooleanVar (généré ou non)
        self.template_vars = {}   # id -> StringVar (chemin modèle, vide = défaut intégré)

        self._build_ui()

    # ------------------------------------------------------------------
    def _default_template(self, etat_id: str) -> str:
        for etat in core.ETATS:
            if etat["id"] == etat_id:
                path = resource_path(os.path.join("resources", etat["resource"]))
                return path if os.path.exists(path) else ""
        return ""

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        header = ttk.Frame(self)
        header.pack(fill="x", **pad)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Importez la Balance N-1 et la Balance N (colonnes Compte, Libellé, "
                 "Débit, Crédit) pour générer automatiquement les états financiers "
                 "sélectionnés ci-dessous.",
            wraplength=740, justify="left",
        ).pack(anchor="w", pady=(4, 0))

        # --- Balances --------------------------------------------------
        frame_n1 = ttk.LabelFrame(self, text="Balance N-1 (exercice précédent)")
        frame_n1.pack(fill="x", **pad)
        self._file_row(frame_n1, self.balance_n1_path, self._choose_balance_n1)

        frame_n = ttk.LabelFrame(self, text="Balance N (exercice en cours)")
        frame_n.pack(fill="x", **pad)
        self._file_row(frame_n, self.balance_n_path, self._choose_balance_n)

        # --- États à générer --------------------------------------------
        frame_etats = ttk.LabelFrame(self, text="États à générer")
        frame_etats.pack(fill="x", **pad)
        for etat in core.ETATS:
            row = ttk.Frame(frame_etats)
            row.pack(fill="x", padx=8, pady=4)

            var = tk.BooleanVar(value=True)
            self.etat_vars[etat["id"]] = var
            ttk.Checkbutton(row, text=etat["label"], variable=var, width=32).pack(side="left")

            tvar = tk.StringVar(value=self._default_template(etat["id"]))
            self.template_vars[etat["id"]] = tvar
            entry = ttk.Entry(row, textvariable=tvar)
            entry.pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(row, text="Modèle…",
                       command=lambda eid=etat["id"]: self._choose_template(eid)).pack(side="left")

        ttk.Label(
            frame_etats,
            text="Un modèle par défaut est déjà intégré pour chaque état. "
                 "Cliquez sur « Modèle… » uniquement pour le remplacer par le vôtre.",
            foreground="#666666",
        ).pack(anchor="w", padx=8, pady=(0, 6))

        # --- Sortie -----------------------------------------------------
        frame_out = ttk.LabelFrame(self, text="Dossier de sortie")
        frame_out.pack(fill="x", **pad)
        self._file_row(frame_out, self.output_dir, self._choose_output_dir)

        # --- Bouton générer ----------------------------------------------
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", **pad)
        self.generate_btn = ttk.Button(btn_frame, text="⚙  Générer les états sélectionnés",
                                        command=self._on_generate)
        self.generate_btn.pack(side="left")
        ttk.Button(btn_frame, text="Ouvrir le dossier de sortie", command=self._open_output_dir).pack(
            side="left", padx=8)

        # --- Journal / résultats ------------------------------------------
        frame_log = ttk.LabelFrame(self, text="Résultat")
        frame_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frame_log, height=14, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    def _file_row(self, parent, var, command):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=6)
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Parcourir…", command=command).pack(side="left", padx=6)

    # ------------------------------------------------------------------
    def _choose_balance_n1(self):
        path = filedialog.askopenfilename(
            title="Choisir la Balance N-1",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.balance_n1_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(os.path.dirname(path), "Etats_financiers"))

    def _choose_balance_n(self):
        path = filedialog.askopenfilename(
            title="Choisir la Balance N",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.balance_n_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(os.path.dirname(path), "Etats_financiers"))

    def _choose_template(self, etat_id: str):
        path = filedialog.askopenfilename(
            title="Choisir le modèle",
            filetypes=[("Classeurs Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.template_vars[etat_id].set(path)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if path:
            self.output_dir.set(path)

    # ------------------------------------------------------------------
    def _log(self, text: str, clear: bool = False):
        self.log_text.configure(state="normal")
        if clear:
            self.log_text.delete("1.0", "end")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_generate(self):
        n1 = self.balance_n1_path.get().strip()
        n = self.balance_n_path.get().strip()
        out_dir = self.output_dir.get().strip()

        if not n1 or not os.path.exists(n1):
            messagebox.showerror(APP_TITLE, "Veuillez sélectionner le fichier de la Balance N-1.")
            return
        if not n or not os.path.exists(n):
            messagebox.showerror(APP_TITLE, "Veuillez sélectionner le fichier de la Balance N.")
            return
        if not out_dir:
            messagebox.showerror(APP_TITLE, "Veuillez indiquer un dossier de sortie.")
            return

        selected_ids = [eid for eid, var in self.etat_vars.items() if var.get()]
        if not selected_ids:
            messagebox.showerror(APP_TITLE, "Sélectionnez au moins un état à générer.")
            return

        templates = {}
        for eid in selected_ids:
            tpath = self.template_vars[eid].get().strip()
            if not tpath or not os.path.exists(tpath):
                messagebox.showerror(APP_TITLE, f"Modèle introuvable pour l'état « {eid} ».")
                return
            templates[eid] = tpath

        self.status_var.set("Génération en cours…")
        self.generate_btn.configure(state="disabled")
        self.update_idletasks()

        try:
            results = core.generate_all_etats(
                balance_n1_path=n1,
                balance_n_path=n,
                output_dir=out_dir,
                templates=templates,
                selected_ids=selected_ids,
            )
        except Exception as e:
            self._log("ERREUR : " + str(e), clear=True)
            self._log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"Échec de la génération :\n{e}")
            self.status_var.set("Échec.")
            self.generate_btn.configure(state="normal")
            return

        self._log(f"Dossier de sortie : {out_dir}", clear=True)
        self._log("")
        any_error = False
        for etat in core.ETATS:
            if etat["id"] not in results:
                continue
            r = results[etat["id"]]
            self._log(f"— {etat['label']} —")
            if isinstance(r, Exception):
                any_error = True
                self._log(f"  ÉCHEC : {r}")
            else:
                self._log(f"  Fichier : {r.output_path}")
                self._log(f"  Cellules calculées : {r.cells_ok}")
                if r.cells_warning:
                    self._log(f"  Avertissements : {len(r.cells_warning)}")
                    for sheet, coord, formula, msg in r.cells_warning:
                        self._log(f"    - [{sheet}!{coord}] {msg}")
                if r.cells_error:
                    any_error = True
                    self._log(f"  Erreurs : {len(r.cells_error)}")
                    for sheet, coord, formula, msg in r.cells_error:
                        self._log(f"    - [{sheet}!{coord}] {formula}\n        -> {msg}")
                else:
                    self._log("  Aucune erreur.")
            self._log("")

        self.status_var.set("Terminé.")
        self.generate_btn.configure(state="normal")
        messagebox.showinfo(
            APP_TITLE,
            "Génération terminée." if not any_error else "Génération terminée, avec des erreurs sur certains états (voir le détail)."
        )

    def _open_output_dir(self):
        out_dir = self.output_dir.get().strip()
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning(APP_TITLE, "Aucun dossier de sortie valide pour l'instant.")
            return
        try:
            os.startfile(out_dir)  # type: ignore[attr-defined]
        except AttributeError:
            webbrowser.open(f"file://{out_dir}")


def main():
    app = EtatsApp()
    app.mainloop()


if __name__ == "__main__":
    main()
