# -*- coding: utf-8 -*-
"""
generer_mot_de_passe.py — UTILITAIRE ADMIN, RÉSERVÉ À VOUS.

Affiche le mot de passe du menu PARAMÈTRES pour le mois en cours et les
mois suivants, à partir de la clé secrète définie dans security.py.

⚠️  NE PAS distribuer ce fichier avec le logiciel (ne l'ajoutez à aucun
    .exe, ne le partagez pas avec vos utilisateurs). Il n'est de toute
    façon jamais inclus dans l'exécutable généré par GitHub Actions,
    puisque main.py ne l'importe pas.

Usage :
    python generer_mot_de_passe.py            (affiche 3 mois : ce mois-ci + 2 suivants)
    python generer_mot_de_passe.py 6           (affiche 6 mois)
"""

import sys
from datetime import datetime

import security

MOIS_FR = [
    "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre",
]


def main():
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 3
    now = datetime.now()

    print("Mots de passe PARAMÈTRES — générés à partir de SECRET_KEY (security.py)")
    print("=" * 70)
    for i in range(n):
        year = now.year + (now.month - 1 + i) // 12
        month = (now.month - 1 + i) % 12 + 1
        when = datetime(year, month, 1)
        pwd = security.generate_monthly_password(when=when)
        label = "  <- CE MOIS-CI" if i == 0 else ""
        print(f"  {MOIS_FR[month - 1].capitalize()} {year} : {pwd}{label}")
    print("=" * 70)
    print("Communiquez uniquement le mot de passe du mois en cours à vos utilisateurs.")


if __name__ == "__main__":
    main()
