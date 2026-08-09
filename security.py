# -*- coding: utf-8 -*-
"""
security.py — Protection par mot de passe mensuel du menu PARAMÈTRES.

Principe : le mot de passe change automatiquement chaque mois, sans accès
réseau, à partir d'une clé secrète (SECRET_KEY) que vous seul connaissez.
Vous n'avez donc pas besoin de redistribuer un nouveau mot de passe chaque
mois "à la main" par un tiers : il suffit de communiquer le mot de passe du
mois à vos utilisateurs (par SMS, appel, etc.), en le calculant vous-même
avec l'utilitaire `generer_mot_de_passe.py` (réservé à vous, non inclus
dans l'exécutable distribué).

IMPORTANT — limite de sécurité à connaître :
Ce mécanisme est un frein pratique contre un utilisateur non technique qui
ouvrirait le logiciel par curiosité — ce n'est PAS un coffre-fort
cryptographique. Le fichier .exe contient ce module (donc la clé secrète)
pour pouvoir vérifier le mot de passe hors-ligne ; une personne disposant de
compétences en rétro-ingénierie pourrait théoriquement l'extraire du binaire
compilé. Pour une protection réellement forte, il faudrait une vérification
côté serveur (hors du périmètre de ce logiciel de bureau autonome).

AVANT DE DIFFUSER LE LOGICIEL : changez SECRET_KEY ci-dessous pour une
valeur que vous seul connaissez, puis recompilez l'exe (push sur GitHub).
"""

import hashlib
import hmac
from datetime import datetime

# Changez cette valeur avant toute diffusion. Gardez-la strictement privée
# (ne la communiquez à personne, ne la publiez jamais, y compris dans un
# dépôt GitHub public — si votre dépôt est public, gardez ce fichier dans
# un dépôt privé, ou changez la clé régulièrement).
SECRET_KEY = "CHANGEZ-MOI-AVANT-DIFFUSION-2026"


def _keystream(length: int, secret: str) -> bytes:
    """Flux de clé déterministe (mode compteur SHA-256) utilisé pour
    chiffrer/déchiffrer les modèles personnalisés stockés sur disque."""
    stream = bytearray()
    counter = 0
    seed = secret.encode("utf-8")
    while len(stream) < length:
        block = hashlib.sha256(seed + counter.to_bytes(4, "big")).digest()
        stream.extend(block)
        counter += 1
    return bytes(stream[:length])


def encrypt_bytes(data: bytes, secret: str = SECRET_KEY) -> bytes:
    """Chiffre `data` (XOR avec un flux de clé dérivé de la clé secrète).

    Ceci n'est PAS un chiffrement de qualité militaire — c'est un
    obfuscateur suffisant pour empêcher un utilisateur non technique
    d'ouvrir le fichier avec Excel ou un éditeur de texte et d'en lire les
    formules. Voir l'avertissement en tête de ce fichier."""
    stream = _keystream(len(data), secret)
    return bytes(b ^ k for b, k in zip(data, stream))


def decrypt_bytes(data: bytes, secret: str = SECRET_KEY) -> bytes:
    """Le XOR est symétrique : déchiffrer = chiffrer avec la même clé."""
    return encrypt_bytes(data, secret)


# Mot de passe Administrateur : fixe (ne change jamais tout seul), connu de
# vous uniquement. Il donne accès à PARAMÈTRES comme le mot de passe
# utilisateur mensuel, ET permet en plus de consulter le mot de passe
# utilisateur du mois directement dans le logiciel (pas besoin de lancer
# generer_mot_de_passe.py). Changez-le si besoin, gardez-le privé.
ADMIN_PASSWORD = "ouaga2001@@@"


def check_admin_password(entered: str) -> bool:
    """Vérifie le mot de passe Administrateur (fixe, sensible à la casse,
    espaces de début/fin ignorés)."""
    if not entered:
        return False
    return entered.strip() == ADMIN_PASSWORD


def check_any_password(entered: str, secret: str = SECRET_KEY, when: datetime = None):
    """Vérifie `entered` contre le mot de passe Administrateur PUIS contre le
    mot de passe utilisateur du mois. Renvoie (autorisé: bool, rôle: str|None)
    où rôle vaut "admin" ou "utilisateur"."""
    if check_admin_password(entered):
        return True, "admin"
    if check_password(entered, secret, when):
        return True, "utilisateur"
    return False, None


def generate_monthly_password(secret: str = SECRET_KEY, when: datetime = None) -> str:
    """Génère le mot de passe du mois de `when` (aujourd'hui par défaut),
    de façon déterministe à partir de la clé secrète. Format lisible :
    XXXX-XXXX (8 caractères hexadécimaux)."""
    when = when or datetime.now()
    period = when.strftime("%Y-%m")
    digest = hmac.new(secret.encode("utf-8"), period.encode("utf-8"), hashlib.sha256).hexdigest().upper()
    return f"{digest[:4]}-{digest[4:8]}"


def check_password(entered: str, secret: str = SECRET_KEY, when: datetime = None) -> bool:
    """Vérifie un mot de passe saisi contre celui du mois en cours (espaces
    et tirets ignorés, insensible à la casse)."""
    if not entered:
        return False
    expected = generate_monthly_password(secret, when)
    norm = lambda s: s.strip().upper().replace(" ", "").replace("-", "")
    return norm(entered) == norm(expected)
