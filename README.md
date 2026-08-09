# États Financiers Automatiques

Logiciel de bureau (Windows, `.exe`) qui génère automatiquement **4 états
financiers** à partir de deux balances comptables (**Balance N-1** et
**Balance N**), en appliquant les modèles Excel fournis :

- **Bilan**
- **Compte de Résultat (SIG)**
- **Situation Financière (FR-BFR-TN)**
- **Flux de Trésorerie (TFT)**

Chaque modèle contient des cellules-formules du type
`CtaCptSoldeDébit("42*")` et des « rubriques » réutilisables
`[xxx.EtLoc]=...`.

## Comment ça marche

1. Vous importez :
   - la **Balance N-1** (fichier Excel ou CSV avec les colonnes `Compte`,
     `Libellé`, `Débit`, `Crédit`)
   - la **Balance N** (même format)
   - vous cochez les états que vous voulez générer (les 4 modèles par
     défaut sont déjà intégrés au logiciel — vous pouvez remplacer chacun
     par le vôtre si besoin)
2. Vous cliquez sur **Générer les états sélectionnés**.
3. Le logiciel calcule chaque cellule-formule de chaque modèle sélectionné
   et enregistre un fichier Excel par état, dans le dossier de sortie
   choisi : `Bilan.xlsx`, `Compte_de_Resultat.xlsx`,
   `Situation_Financiere.xlsx`, `Flux_de_Tresorerie.xlsx`.

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

Une formule non reconnue (référence externe non gérée, fonction inconnue —
ex. la fonction `Ratio(...)` utilisée une fois dans le modèle "Situation
Financière" fourni, qui n'appartient pas au langage `CtaCptSolde...`) est
signalée dans le journal de résultat sans bloquer le reste du calcul — la
cellule correspondante affiche `#ERREUR`.

Si une rubrique est référencée mais n'est **jamais définie nulle part**
dans le modèle (trou dans le modèle, ex. une ligne supprimée par erreur),
elle est traitée comme **0** plutôt que de bloquer toute la chaîne de
calcul qui en dépend — un avertissement le signale dans le journal, sans
empêcher la génération du reste de l'état.

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
   téléchargez l'exécutable dans l'artifact **EtatsFinanciers-windows**.
3. Pour publier une **release** téléchargeable directement (avec un lien
   fixe), créez un tag de version :
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
   Le workflow attachera alors `EtatsFinanciers.exe` à la release GitHub
   correspondante.

## Compilation locale (si vous avez un PC Windows sous la main)

```bash
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name "EtatsFinanciers" --add-data "resources;resources" main.py
```
L'exécutable est généré dans `dist/EtatsFinanciers.exe`.

## Structure du projet

```
bilan-auto/
├── main.py                              interface graphique (Tkinter)
├── core.py                              moteur de calcul (formules CtaCptSolde...)
├── resources/
│   ├── modele_bilan.xlsx                modèle Bilan par défaut
│   ├── modele_resultat.xlsx             modèle Compte de Résultat (SIG) par défaut
│   ├── modele_situation.xlsx            modèle Situation Financière (FR-BFR-TN) par défaut
│   └── modele_flux.xlsx                 modèle Flux de Trésorerie (TFT) par défaut
├── requirements.txt
├── .github/workflows/build.yml          build automatique de l'exe (GitHub Actions)
└── README.md
```

## Protection par mot de passe — formules cachées aux utilisateurs

Le menu **PARAMÈTRES** (et le bouton « Modèle externe… ») est protégé par
mot de passe. **Deux niveaux d'accès** :

| Rôle | Mot de passe | Donne accès à |
|---|---|---|
| **Utilisateur** | change automatiquement chaque mois (ex. `F1B9-2E41` en août 2026) | Ouvrir/modifier les 4 modèles via PARAMÈTRES |
| **Administrateur** | fixe, défini par vous (`ADMIN_PASSWORD` dans `security.py`) | Tout ce que fait l'utilisateur, **+** consulter directement dans le logiciel le mot de passe utilisateur du mois (menu *PARAMÈTRES → 🔑 Mot de passe utilisateur du mois*) |

Concrètement : vous vous connectez avec le mot de passe Administrateur, et
le logiciel vous affiche lui-même le mot de passe Utilisateur du mois (et
des 2 mois suivants) — plus besoin de lancer `generer_mot_de_passe.py` à la
main. Vous communiquez ensuite ce mot de passe Utilisateur à vos
utilisateurs (SMS, appel...), qui pourront l'utiliser pour ouvrir
PARAMÈTRES ce mois-ci — le mois suivant, ce mot de passe change
automatiquement.

### Avant de diffuser le logiciel

1. Ouvrez `security.py` :
   - Changez `SECRET_KEY` (détermine le mot de passe Utilisateur mensuel)
   - Changez `ADMIN_PASSWORD` si vous voulez un mot de passe Administrateur
     différent de la valeur actuelle
2. Gardez ces deux valeurs strictement privées (si votre dépôt GitHub est
   public, gardez ce fichier dans un dépôt **privé**, ou changez les
   valeurs très régulièrement).
3. Recompilez l'exe (poussez sur GitHub — voir plus haut).

### Consulter le mot de passe utilisateur sans ouvrir le logiciel

Vous pouvez aussi, sur votre poste, exécuter :
```bash
python generer_mot_de_passe.py
```
Ce script n'est jamais inclus dans le `.exe` généré (il n'est pas importé
par `main.py`), donc vos utilisateurs n'y ont pas accès même s'ils
examinent les fichiers du logiciel.

### Modèles personnalisés chiffrés sur disque

Si vous modifiez un modèle via PARAMÈTRES, il est enregistré **chiffré**
dans un dossier `modeles_personnalises/` créé à côté de l'exécutable — pas
en `.xlsx` en clair. Un utilisateur qui ouvrirait ce dossier dans
l'Explorateur Windows ne trouvera que des fichiers `.dat` illisibles, pas
ouvrables dans Excel.

### Limite honnête de cette protection

Ce mécanisme est un **frein pratique** contre un utilisateur non technique
curieux — ce n'est pas un coffre-fort cryptographique de niveau
professionnel. Le fichier `.exe` contient nécessairement `security.py`
(donc `SECRET_KEY`) pour pouvoir vérifier le mot de passe hors-ligne ; une
personne avec des compétences en rétro-ingénierie de binaires Python
pourrait théoriquement l'extraire. Pour une protection réellement
inviolable, il faudrait une vérification côté serveur, ce qui dépasse le
cadre d'un logiciel de bureau autonome. Pour l'usage décrit (empêcher des
utilisateurs classiques de voir les formules), c'est largement suffisant.

## Menu LIASSE FISCALE — fiche d'identification

Le menu **LIASSE FISCALE** (accessible sans mot de passe — ce ne sont pas
des formules à protéger) ouvre une page de saisie pour les 7 informations
communes à toutes les feuilles de la liasse fiscale complète :

- Dénomination sociale de l'entité
- Adresse
- N° IFU du contribuable
- N° de télédéclarant (NES)
- Sigle usuel
- Durée (en mois)
- Exercice clos le

Dans le modèle de liasse fiscale (`resources/modele_liasse_fiscale.xlsx`),
ces informations ne sont **pas ressaisies sur chaque feuille** : presque
toutes les ~80 feuilles contiennent une formule qui pointe, directement ou
par une chaîne de renvois, vers une poignée de cellules racines (feuille
`GARDE`, principalement). Renseigner ces 7 champs une seule fois dans cette
page suffit donc à les faire apparaître partout dans la liasse.

- **💾 Enregistrer** conserve les valeurs saisies (fichier
  `modeles_personnalises/liasse_fiscale_identite.json`, persistant d'un
  lancement à l'autre — modifiable à tout moment).
- **📄 Exporter le classeur…** applique ces informations au modèle de
  liasse fiscale et enregistre le résultat à l'emplacement de votre choix.

**Portée actuelle** : cette page couvre la fiche d'identification, **et
génère aussi automatiquement le Bilan, le Compte de Résultat et le Tableau
des Flux de Trésorerie de la liasse complète** à partir de vos balances
(bouton « ⚙ Générer la Liasse Fiscale complète… »), en réutilisant les
comptes SYSCOHADA déjà validés dans les 4 petits modèles.

**Résultats de validation** (données réelles fournies) :
- **BILAN** : équilibre Actif = Passif **parfait** (écart nul), et total
  identique au petit modèle "Bilan Synthétique" déjà validé séparément.
- **RESULTAT** : Résultat Net Comptable identique à la décimale près au
  Résultat calculé dans le Bilan.
- **TFT** : structure et sommes internes cohérentes (zéro erreur de
  formule) ; n'a pas pu être validé numériquement avec l'échantillon
  fourni car ses deux feuilles de balance (`Balance` et `Balance 2023 (2)`)
  s'y sont révélées strictement identiques (donc aucune variation N/N-1
  réelle à mesurer) — cela fonctionnera normalement avec de vraies
  données N-1 distinctes.
- **NOTE 34** (Situation Financière — indicateurs, CAFG, FR-BFR-TN) : cette
  feuille du modèle officiel contenait déjà de vraies formules Excel
  pointant vers BILAN/RESULTAT/TFT (ex. ratios de rentabilité) ; les
  cellules manquantes ont été complétées de la même façon (références
  Excel directes vers les 3 autres feuilles calculées, plus quelques
  comptes CtaCptSolde pour la chaîne de la CAFG). Cohérence vérifiée :
  Trésorerie nette = Contrôle trésorerie à l'euro près, Variation de
  trésorerie = TFT à l'euro près.

**Limite à connaître** : le mapping comptable REF-code → comptes SYSCOHADA
pour ces 3 états est un travail de reconstruction (le modèle officiel
fourni ne contenait aucune formule d'origine, contrairement aux 4 petits
modèles). La majorité des lignes réutilisent des comptes déjà validés ;
quelques lignes de détail (ex. répartition fine de certaines
immobilisations incorporelles, quelques provisions) restent estimées et
méritent une relecture avant tout usage officiel — voir
`liasse_build/confiance.json` dans l'historique de développement pour le
détail ligne par ligne (« validé » vs « estimé » vs « absent »). Les
autres notes annexes détaillées de la liasse (hors NOTE 34) ne sont pas
couvertes par ce calcul automatique (seuls BILAN, RESULTAT, TFT et
NOTE 34/Situation Financière le sont).

## Faire évoluer le modèle de Bilan

Le modèle n'est pas codé en dur : le moteur lit **n'importe quelle**
cellule commençant par `=` et contenant une des fonctions `CtaCptSolde...`,
où qu'elle se trouve dans la feuille. Pour ajouter une ligne à un état
(ex. immobilisations, capitaux propres), il suffit d'ouvrir le modèle dans
Excel, d'ajouter la ligne avec le libellé et la formule voulue, en
respectant la syntaxe ci-dessus — aucune modification du code n'est
nécessaire.

## Menu PARAMÈTRES — modifier les modèles directement dans le logiciel

Pas besoin d'Excel : le menu **PARAMÈTRES**, en haut de la fenêtre, liste
les 4 états simples **et la Liasse Fiscale** (sous-menu séparé, avec accès
direct aux 3 feuilles calculées : BILAN, RESULTAT, TFT). Cliquer sur l'un
d'eux ouvre une grille (comme un mini-tableur) affichant toutes les
cellules du modèle — libellés et formules, colonnes année N **et** année
N-1 comprises.

- Modifiez n'importe quelle cellule (un libellé, ou une formule comme
  `=CtaCptSoldeDébit("42*")`, `=CtaCptSoldeDébitNm1("42*")` pour le N-1, ou
  une rubrique `=[011.EtLoc]=...`).
- **Enregistrer** sauvegarde le modèle dans un dossier `modeles_personnalises/`
  créé à côté de l'exécutable — vos modifications sont donc conservées
  d'un lancement à l'autre du logiciel, et utilisées automatiquement à la
  prochaine génération.
- **Restaurer le modèle d'origine** annule vos modifications et revient au
  modèle intégré par défaut pour cet état.

Le champ « Modèle… » de l'écran principal reste disponible pour charger
ponctuellement un fichier externe différent, sans passer par PARAMÈTRES.
