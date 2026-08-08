# -*- coding: utf-8 -*-
"""
main.py — Bilan Automatique
Interface graphique : importer Balance N-1 + Balance N, générer le Bilan
à partir du modèle Excel, en appliquant les formules CtaCptSolde...
"""

import os
import sys
import traceback
import webbrowser
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import core

APP_TITLE = "Bilan Automatique"
APP_VERSION = "1.0"


def resource_path(relative_path: str) -> str:
    """Résout le chemin d'une ressource, y compris depuis un .exe PyInstaller
    (onefile) où les fichiers sont extraits dans un dossier temporaire."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


DEFAULT_TEMPLATE = resource_path(os.path.join("resources", "modele_bilan.xlsx"))


class BilanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} — v{APP_VERSION}")
        self.geometry("720x520")
        self.minsize(660, 480)

        self.template_path = tk.StringVar(value=DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "")
        self.balance_n1_path = tk.StringVar()
        self.balance_n_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Prêt.")

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        header = ttk.Frame(self)
        header.pack(fill="x", **pad)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Importez la Balance N-1 et la Balance N (fichiers Excel/CSV : "
                 "colonnes Compte, Libellé, Débit, Crédit) pour générer "
                 "automatiquement le Bilan selon le modèle.",
            wraplength=680, justify="left",
        ).pack(anchor="w", pady=(4, 0))

        # --- Modèle -------------------------------------------------
        frame_tpl = ttk.LabelFrame(self, text="Modèle de Bilan (feuille 'BILAN' avec formules)")
        frame_tpl.pack(fill="x", **pad)
        self._file_row(frame_tpl, self.template_path, self._choose_template,
                        "Utiliser le modèle intégré par défaut, ou choisir un autre fichier .xlsx")

        # --- Balance N-1 ----------------------------------------------
        frame_n1 = ttk.LabelFrame(self, text="Balance N-1 (exercice précédent)")
        frame_n1.pack(fill="x", **pad)
        self._file_row(frame_n1, self.balance_n1_path, self._choose_balance_n1)

        # --- Balance N --------------------------------------------------
        frame_n = ttk.LabelFrame(self, text="Balance N (exercice en cours)")
        frame_n.pack(fill="x", **pad)
        self._file_row(frame_n, self.balance_n_path, self._choose_balance_n)

        # --- Sortie -----------------------------------------------------
        frame_out = ttk.LabelFrame(self, text="Fichier de sortie")
        frame_out.pack(fill="x", **pad)
        self._file_row(frame_out, self.output_path, self._choose_output, save_mode=True)

        # --- Bouton générer ----------------------------------------------
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", **pad)
        self.generate_btn = ttk.Button(btn_frame, text="⚙  Générer le Bilan", command=self._on_generate)
        self.generate_btn.pack(side="left")
        ttk.Button(btn_frame, text="Ouvrir le fichier généré", command=self._open_output).pack(side="left", padx=8)

        # --- Journal / résultats ------------------------------------------
        frame_log = ttk.LabelFrame(self, text="Résultat")
        frame_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frame_log, height=12, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    def _file_row(self, parent, var, command, hint=None, save_mode=False):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=6)
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Parcourir…", command=command).pack(side="left", padx=6)
        if hint:
            ttk.Label(parent, text=hint, foreground="#666666").pack(anchor="w", padx=8, pady=(0, 4))

    # ------------------------------------------------------------------
    def _choose_template(self):
        path = filedialog.askopenfilename(
            title="Choisir le modèle de Bilan",
            filetypes=[("Classeurs Excel", "*.xlsx"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.template_path.set(path)

    def _choose_balance_n1(self):
        path = filedialog.askopenfilename(
            title="Choisir la Balance N-1",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.balance_n1_path.set(path)
            if not self.output_path.get():
                self._suggest_output()

    def _choose_balance_n(self):
        path = filedialog.askopenfilename(
            title="Choisir la Balance N",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.balance_n_path.set(path)
            if not self.output_path.get():
                self._suggest_output()

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Enregistrer le Bilan sous…",
            defaultextension=".xlsx",
            filetypes=[("Classeur Excel", "*.xlsx")],
            initialfile="Bilan.xlsx",
        )
        if path:
            self.output_path.set(path)

    def _suggest_output(self):
        base_dir = os.path.dirname(self.balance_n_path.get() or self.balance_n1_path.get() or os.getcwd())
        self.output_path.set(os.path.join(base_dir, "Bilan_genere.xlsx"))

    # ------------------------------------------------------------------
    def _log(self, text: str, clear: bool = False):
        self.log_text.configure(state="normal")
        if clear:
            self.log_text.delete("1.0", "end")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_generate(self):
        template = self.template_path.get().strip()
        n1 = self.balance_n1_path.get().strip()
        n = self.balance_n_path.get().strip()
        out = self.output_path.get().strip()

        if not template or not os.path.exists(template):
            messagebox.showerror(APP_TITLE, "Veuillez sélectionner un modèle de Bilan valide (.xlsx).")
            return
        if not n1 or not os.path.exists(n1):
            messagebox.showerror(APP_TITLE, "Veuillez sélectionner le fichier de la Balance N-1.")
            return
        if not n or not os.path.exists(n):
            messagebox.showerror(APP_TITLE, "Veuillez sélectionner le fichier de la Balance N.")
            return
        if not out:
            messagebox.showerror(APP_TITLE, "Veuillez indiquer le fichier de sortie.")
            return

        self.status_var.set("Génération en cours…")
        self.generate_btn.configure(state="disabled")
        self.update_idletasks()

        try:
            report = core.generate_bilan(
                template_path=template,
                balance_n1_path=n1,
                balance_n_path=n,
                output_path=out,
            )
        except Exception as e:
            self._log("ERREUR : " + str(e), clear=True)
            self._log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"Échec de la génération :\n{e}")
            self.status_var.set("Échec.")
            self.generate_btn.configure(state="normal")
            return

        self._log(f"Bilan généré : {report.output_path}", clear=True)
        self._log(f"Cellules calculées avec succès : {report.cells_ok}")
        if report.cells_error:
            self._log(f"Cellules en erreur : {len(report.cells_error)}")
            for sheet, coord, formula, msg in report.cells_error:
                self._log(f"  - [{sheet}!{coord}] {formula}\n      -> {msg}")
        else:
            self._log("Aucune erreur : toutes les formules ont été calculées.")

        self.status_var.set("Terminé.")
        self.generate_btn.configure(state="normal")
        messagebox.showinfo(APP_TITLE, "Le Bilan a été généré avec succès." if not report.cells_error
                             else "Bilan généré, avec quelques cellules en erreur (voir le détail).")

    def _open_output(self):
        out = self.output_path.get().strip()
        if not out or not os.path.exists(out):
            messagebox.showwarning(APP_TITLE, "Aucun fichier généré à ouvrir pour l'instant.")
            return
        try:
            os.startfile(out)  # type: ignore[attr-defined]
        except AttributeError:
            webbrowser.open(f"file://{out}")


def main():
    app = BilanApp()
    app.mainloop()


if __name__ == "__main__":
    main()
