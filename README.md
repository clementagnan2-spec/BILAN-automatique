# Bilan Automatique

Logiciel de bureau (Windows, `.exe`) qui génère automatiquement le **Bilan**
à partir de deux balances comptables (**Balance N-1** et **Balance N**),
en appliquant le modèle Excel fourni (feuille `BILAN` avec des formules du
type `CtaCptSoldeDébit("42*")`).

## Comment ça marche

1. Vous importez :
   - un **modèle** de Bilan (`.xlsx`) contenant une feuille `BILAN` avec des
     cellules-formules `CtaCptSoldeDébit(...)`, `CtaCptSoldeCrédit(...)`,
     `CtaCptSoldeDébitNm1(...)`, `CtaCptSoldeCréditNm1(...)`,
     `CtaCptSolde(...)`, `CtaCptSoldeNm1(...)`. Un modèle par défaut est déjà
     intégré au logiciel (basé sur `resources/modele_bilan.xlsx`) — vous
     pouvez le remplacer par le vôtre si votre présentation évolue.
   - la **Balance N-1** (fichier Excel ou CSV avec les colonnes `Compte`,
     `Libellé`, `Débit`, `Crédit`)
   - la **Balance N** (même format)
2. Vous cliquez sur **Générer le Bilan**.
3. Le logiciel calcule chaque cellule-formule du modèle et enregistre un
   nouveau classeur Excel avec les montants calculés, prêt à l'emploi.

## Logique des formules

| Formule | Signification |
|---|---|
| `CtaCptSoldeDébit("42*")` | Somme des soldes (Balance N) de tous les comptes de racine **42** dont le solde est **débiteur** |
| `CtaCptSoldeCrédit("42*")` | Idem mais comptes **créditeurs** (Balance N) |
| `CtaCptSoldeDébitNm1("41*")` | Comme `CtaCptSoldeDébit`, mais sur la **Balance N-1** |
| `CtaCptSoldeCréditNm1("41*")` | Comme `CtaCptSoldeCrédit`, mais sur la **Balance N-1** |
| `CtaCptSoldeDébit("50*","56*")` | Avec **deux racines**, ce n'est pas une union mais une **plage** : tous les comptes dont la racine est comprise entre 50 et 56 inclus |
| `CtaCptSolde("280*","2869*")` | Solde **net** (débit − crédit, sans filtrage de sens), utile pour soustraire des amortissements/provisions |
| `CtaCptSoldeNm1(...)` | Idem, sur la Balance N-1 |
| `[Rxxx.EtLoc]=<formule>` | Marque la cellule comme "rubrique Rxxx" réutilisable ailleurs (ex. pour un TOTAL qui additionne plusieurs sous-totaux déjà calculés) |
| `[Rxxx.EtLoc]` (dans une autre formule) | Réutilise la valeur déjà calculée de la rubrique Rxxx (ex. `TOTAL I = [R120.EtLoc]+[R130.EtLoc]+...`) |

Le solde d'un compte est calculé comme **Débit − Crédit** sur les lignes de
la balance importée (les lignes en double pour un même compte sont
additionnées automatiquement).

Le moteur résout les formules en **plusieurs passes** : une cellule qui
référence une rubrique `[Rxxx.EtLoc]` pas encore calculée est simplement
recalculée à la passe suivante, jusqu'à convergence — l'ordre des lignes
dans le modèle n'a donc pas besoin d'être parfait.

Une formule non reconnue (référence externe non gérée, fonction inconnue)
est signalée dans le journal de résultat sans bloquer le reste du calcul —
la cellule correspondante affiche `#ERREUR`.

## Formats de modèle acceptés

Le modèle de Bilan peut être fourni en :
- **`.xlsx`** natif (recommandé)
- **`.xls`** au format XML "Excel 2003 Spreadsheet" (export courant depuis
  certains logiciels comptables) — le logiciel le convertit automatiquement,
  sans dépendance externe à installer.

Le nom de la feuille contenant les formules n'a pas besoin de s'appeler
`BILAN` : si cette feuille n'existe pas, le logiciel utilise automatiquement
la feuille qui contient le plus de formules `CtaCptSolde...`.

## Utilisation en local (sans passer par l'exe)

```bash
pip install -r requirements.txt
python main.py
```

## Compilation de l'exécutable Windows (.exe)

La compilation d'un `.exe` nécessite Windows (PyInstaller ne fait pas de
compilation croisée). Ce dépôt contient un **workflow GitHub Actions**
(`.github/workflows/build.yml`) qui s'en charge automatiquement :

1. Créez un dépôt GitHub et poussez ce projet :
   ```bash
   git init
   git add .
   git commit -m "Initial commit — Bilan Automatique"
   git branch -M main
   git remote add origin https://github.com/<votre-compte>/<votre-repo>.git
   git push -u origin main
   ```
2. Dans l'onglet **Actions** du dépôt GitHub, le workflow *Build Windows EXE*
   se déclenche automatiquement. Une fois terminé (quelques minutes),
   téléchargez l'exécutable dans l'artifact **BilanAutomatique-windows**.
3. Pour publier une **release** téléchargeable directement (avec un lien
   fixe), créez un tag de version :
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
   Le workflow attachera alors `BilanAutomatique.exe` à la release GitHub
   correspondante.

## Compilation locale (si vous avez un PC Windows sous la main)

```bash
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name "BilanAutomatique" --add-data "resources;resources" main.py
```
L'exécutable est généré dans `dist/BilanAutomatique.exe`.

## Structure du projet

```
bilan-auto/
├── main.py                        interface graphique (Tkinter)
├── core.py                        moteur de calcul (formules CtaCptSolde...)
├── resources/modele_bilan.xlsx    modèle de Bilan par défaut
├── requirements.txt
├── .github/workflows/build.yml    build automatique de l'exe (GitHub Actions)
└── README.md
```

## Faire évoluer le modèle de Bilan

Le modèle n'est pas codé en dur : le moteur lit **n'importe quelle**
cellule commençant par `=` et contenant une des fonctions `CtaCptSolde...`,
où qu'elle se trouve dans la feuille `BILAN`. Pour ajouter une ligne au
Bilan (ex. immobilisations, capitaux propres), il suffit d'ouvrir
`resources/modele_bilan.xlsx` dans Excel, d'ajouter la ligne avec le
libellé et la formule voulue, en respectant la syntaxe ci-dessus — aucune
modification du code n'est nécessaire.
