# -*- coding: utf-8 -*-
"""
core.py — Moteur de génération automatique du Bilan.

Principe :
1. On charge deux "balances" (Balance N-1 et Balance N), chacune une liste
   de lignes (Compte, Libellé, Débit, Crédit).
2. On calcule, pour chaque compte, son solde net = Débit - Crédit (les
   valeurs vides / non numériques comptent pour 0).
3. Le modèle du Bilan (feuille "BILAN" d'un classeur Excel) contient des
   cellules-formules utilisant des pseudo-fonctions inspirées du langage
   de l'utilisateur :

     CtaCptSoldeDébit("42*")            -> somme des comptes de racine 42
                                            dont le solde (Balance N) est
                                            débiteur (> 0)
     CtaCptSoldeCrédit("42*")           -> idem mais comptes créditeurs (<0),
                                            valeur renvoyée positive
     CtaCptSoldeDébitNm1("41*")         -> comme CtaCptSoldeDébit mais sur
                                            la Balance N-1
     CtaCptSoldeCréditNm1("41*")        -> comme CtaCptSoldeCrédit mais sur
                                            la Balance N-1
     CtaCptSolde("280*","2869*")        -> solde NET (peut être négatif),
                                            sans filtrage de sens, sur N
     CtaCptSoldeNm1(...)                -> idem sur N-1

   Avec DEUX arguments, la fonction ne fait pas l'union des deux racines
   mais la PLAGE (intervalle) entre les deux racines, bornes incluses :
   CtaCptSoldeDébit("50*","56*") = comptes dont la racine est comprise
   entre 50 et 56 inclus (50,51,...,56), pas seulement 50 et 56.

   Ces cellules-formules ne sont PAS de vraies formules Excel (Excel ne
   connaît pas ces fonctions) : ce sont de simples chaînes de texte que ce
   moteur lit, interprète et remplace par leur valeur calculée.

4. Le résultat est écrit dans une copie du classeur modèle (mêmes libellés,
   même mise en forme), avec les cellules-formules remplacées par les
   montants calculés.
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd


# --------------------------------------------------------------------------
# 1. Chargement des balances
# --------------------------------------------------------------------------

REQUIRED_COLS = {
    "compte": ["compte", "compte général", "compte general", "n° compte", "numero compte", "numéro compte"],
    "libelle": ["libellé", "libelle", "intitulé", "intitule", "designation", "désignation"],
    "debit": ["débit", "debit", "mvt débit", "mvt debit", "solde débit", "solde debit"],
    "credit": ["crédit", "credit", "mvt crédit", "mvt credit", "solde crédit", "solde credit"],
}


def _to_float(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        try:
            if v != v:  # NaN
                return 0.0
        except Exception:
            pass
        return float(v)
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    s = s.replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _find_header_row(raw: pd.DataFrame) -> int:
    """Cherche la ligne d'en-tête (celle qui contient 'compte')."""
    for i in range(min(10, len(raw))):
        row_vals = [str(x).strip().lower() for x in raw.iloc[i].tolist()]
        if any(v in REQUIRED_COLS["compte"] for v in row_vals):
            return i
    return 0


def _map_columns(columns) -> dict:
    mapping = {}
    norm_cols = {str(c).strip().lower(): c for c in columns}
    for key, aliases in REQUIRED_COLS.items():
        for alias in aliases:
            if alias in norm_cols:
                mapping[key] = norm_cols[alias]
                break
    return mapping


@dataclass
class Balance:
    """Balance comptable agrégée : {compte(str) -> (debit, credit)}."""
    soldes: dict = field(default_factory=dict)   # compte -> solde net (debit-credit)
    libelles: dict = field(default_factory=dict)  # compte -> libellé (premier rencontré)

    def solde(self, compte: str) -> float:
        return self.soldes.get(compte, 0.0)


def load_balance(path: str, sheet_name: Optional[str] = None) -> Balance:
    """Charge un fichier de balance (xlsx/xls/csv) et renvoie un objet Balance
    avec un solde net (débit-crédit) par compte (les lignes répétées pour un
    même compte sont additionnées)."""
    path = str(path)
    if path.lower().endswith(".csv"):
        raw = pd.read_csv(path, header=None, dtype=object)
    else:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        if isinstance(raw, dict):
            # plusieurs feuilles retournées -> prendre la première non vide
            raw = next(iter(raw.values()))

    header_row = _find_header_row(raw)
    header = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].copy()
    data.columns = header

    colmap = _map_columns(data.columns)
    missing = [k for k in ("compte", "debit", "credit") if k not in colmap]
    if missing:
        raise ValueError(
            "Colonnes introuvables dans le fichier de balance : %s. "
            "Colonnes attendues : Compte, Libellé, Débit, Crédit." % ", ".join(missing)
        )

    bal = Balance()
    for _, row in data.iterrows():
        compte_raw = row.get(colmap["compte"])
        if compte_raw is None or (isinstance(compte_raw, float) and compte_raw != compte_raw):
            continue
        compte = str(compte_raw).strip()
        if compte == "" or compte.lower() == "nan":
            continue
        # normaliser un compte du type "401100.0" -> "401100"
        if compte.endswith(".0"):
            compte = compte[:-2]

        debit = _to_float(row.get(colmap["debit"]))
        credit = _to_float(row.get(colmap["credit"]))
        libelle = str(row.get(colmap.get("libelle"), "") or "")

        bal.soldes[compte] = bal.soldes.get(compte, 0.0) + (debit - credit)
        if compte not in bal.libelles and libelle and libelle.lower() != "nan":
            bal.libelles[compte] = libelle

    return bal


# --------------------------------------------------------------------------
# 2. Fonctions de calcul CtaCptSolde...
# --------------------------------------------------------------------------

def _racine_range(prefixes: tuple[str, ...]) -> tuple[str, str]:
    """Normalise les préfixes (avec ou sans '*') et renvoie (debut, fin)."""
    clean = [p.rstrip("*") for p in prefixes]
    if len(clean) == 1:
        return clean[0], clean[0]
    return clean[0], clean[1]


def _compte_dans_plage(compte: str, debut: str, fin: str) -> bool:
    """Teste si `compte` a une racine comprise entre `debut` et `fin`
    (bornes incluses), plage définie sur le nombre de chiffres du préfixe
    le plus long parmi les deux bornes."""
    if debut == fin:
        return compte.startswith(debut)
    L = max(len(debut), len(fin))
    borne_min = int(debut.ljust(L, "0"))
    borne_max = int(fin.ljust(L, "9"))
    prefix_compte = compte[:L]
    if len(prefix_compte) < L:
        prefix_compte = prefix_compte.ljust(L, "0")
    try:
        val = int(prefix_compte)
    except ValueError:
        return False
    return borne_min <= val <= borne_max


def _comptes_matching(balance: Balance, prefixes: tuple[str, ...]):
    debut, fin = _racine_range(prefixes)
    for compte in balance.soldes:
        if _compte_dans_plage(compte, debut, fin):
            yield compte


def cta_cpt_solde_debit(balance: Balance, *prefixes: str) -> float:
    """Somme des soldes (positifs) des comptes débiteurs matchant les racines."""
    total = 0.0
    for compte in _comptes_matching(balance, prefixes):
        s = balance.solde(compte)
        if s > 0:
            total += s
    return total


def cta_cpt_solde_credit(balance: Balance, *prefixes: str) -> float:
    """Somme des soldes (valeur absolue) des comptes créditeurs matchant les racines."""
    total = 0.0
    for compte in _comptes_matching(balance, prefixes):
        s = balance.solde(compte)
        if s < 0:
            total += -s
    return total


def cta_cpt_solde(balance: Balance, *prefixes: str) -> float:
    """Solde net (peut être positif ou négatif), sans filtrage de sens."""
    total = 0.0
    for compte in _comptes_matching(balance, prefixes):
        total += balance.solde(compte)
    return total


# --------------------------------------------------------------------------
# 3. Évaluateur de formules
# --------------------------------------------------------------------------

# Ordre important : les noms longs (...Nm1) doivent être testés avant les
# noms courts pour ne pas être coupés par erreur lors du remplacement.
_FUNC_TOKENS = [
    ("CtaCptSoldeDébitNm1", "__F_DEBIT_NM1__"),
    ("CtaCptSoldeDebitNm1", "__F_DEBIT_NM1__"),
    ("CtaCptSoldeCréditNm1", "__F_CREDIT_NM1__"),
    ("CtaCptSoldeCreditNm1", "__F_CREDIT_NM1__"),
    ("CtaCptSoldeNm1", "__F_SOLDE_NM1__"),
    ("CtaCptSoldeDébit", "__F_DEBIT__"),
    ("CtaCptSoldeDebit", "__F_DEBIT__"),
    ("CtaCptSoldeCrédit", "__F_CREDIT__"),
    ("CtaCptSoldeCredit", "__F_CREDIT__"),
    ("CtaCptSolde", "__F_SOLDE__"),
]


class FormulaError(Exception):
    pass


def is_formula(value) -> bool:
    return isinstance(value, str) and value.strip().startswith("=") and "CtaCptSolde" in value


def evaluate_formula(formula: str, balance_n: Balance, balance_n1: Balance):
    """Évalue une formule du modèle de Bilan (chaîne commençant par '=')
    et renvoie la valeur numérique calculée.

    Les formules non reconnues (ex : références externes du type
    [R200.EtLoc] laissées par un ancien modèle de liasse) lèvent une
    FormulaError plutôt que de faire planter tout le traitement."""
    expr = formula.strip()
    if expr.startswith("="):
        expr = expr[1:]

    if "[" in expr or "]" in expr:
        raise FormulaError("Référence externe non prise en charge : %s" % formula)

    py_expr = expr
    for name, token in _FUNC_TOKENS:
        py_expr = py_expr.replace(name, token)

    if "CtaCptSolde" in py_expr:
        raise FormulaError("Fonction non reconnue : %s" % formula)
    # il ne doit rester aucune lettre en dehors de nos tokens internes __F_..__
    leftover = re.sub(r"__F_[A-Z0-9_]+__", "", py_expr)
    if re.search(r"[A-Za-zÀ-ÿ]", leftover):
        raise FormulaError("Fonction ou référence non reconnue : %s" % formula)

    def F_debit(*prefixes):
        return cta_cpt_solde_debit(balance_n, *prefixes)

    def F_credit(*prefixes):
        return cta_cpt_solde_credit(balance_n, *prefixes)

    def F_solde(*prefixes):
        return cta_cpt_solde(balance_n, *prefixes)

    def F_debit_nm1(*prefixes):
        return cta_cpt_solde_debit(balance_n1, *prefixes)

    def F_credit_nm1(*prefixes):
        return cta_cpt_solde_credit(balance_n1, *prefixes)

    def F_solde_nm1(*prefixes):
        return cta_cpt_solde(balance_n1, *prefixes)

    safe_globals = {"__builtins__": {}}
    safe_locals = {
        "__F_DEBIT__": F_debit,
        "__F_CREDIT__": F_credit,
        "__F_SOLDE__": F_solde,
        "__F_DEBIT_NM1__": F_debit_nm1,
        "__F_CREDIT_NM1__": F_credit_nm1,
        "__F_SOLDE_NM1__": F_solde_nm1,
    }

    try:
        return eval(py_expr, safe_globals, safe_locals)  # noqa: S307 (namespace restreint)
    except Exception as e:
        raise FormulaError("Erreur d'évaluation (%s) : %s" % (formula, e))


# --------------------------------------------------------------------------
# 4. Génération du classeur Bilan à partir d'un modèle
# --------------------------------------------------------------------------

@dataclass
class GenerationReport:
    cells_ok: int = 0
    cells_error: list = field(default_factory=list)  # [(sheet, coord, formula, message)]
    output_path: str = ""


def generate_bilan(template_path: str, balance_n1_path: str, balance_n_path: str,
                    output_path: str,
                    sheet_n1: str = None, sheet_n: str = None,
                    bilan_sheet: str = "BILAN") -> GenerationReport:
    """Génère le fichier Bilan à partir du modèle et des deux balances.

    - template_path : classeur .xlsx contenant la feuille "BILAN" avec les
      formules CtaCptSolde... (les autres feuilles du modèle, s'il y en a,
      sont recopiées telles quelles).
    - balance_n1_path / balance_n_path : fichiers de balance (xlsx/csv) à
      importer, avec colonnes Compte / Libellé / Débit / Crédit.
    - output_path : fichier .xlsx de sortie.
    """
    balance_n1 = load_balance(balance_n1_path, sheet_name=sheet_n1)
    balance_n = load_balance(balance_n_path, sheet_name=sheet_n)

    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path, data_only=False)

    if bilan_sheet not in wb.sheetnames:
        raise ValueError("La feuille '%s' est introuvable dans le modèle." % bilan_sheet)

    report = GenerationReport(output_path=output_path)
    ws = wb[bilan_sheet]

    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if is_formula(val):
                try:
                    result = evaluate_formula(val, balance_n, balance_n1)
                    cell.value = result
                    report.cells_ok += 1
                except FormulaError as e:
                    report.cells_error.append((bilan_sheet, cell.coordinate, val, str(e)))
                    cell.value = "#ERREUR"

    # On remplace aussi, par commodité, les feuilles BALANCE par les
    # nouvelles données importées (traçabilité), si ces feuilles existent
    # dans le modèle et si on veut les régénérer -> on laisse le modèle
    # inchangé sur ce point pour ne pas dénaturer un modèle personnalisé ;
    # seule la feuille BILAN est recalculée.

    wb.save(output_path)
    return report
