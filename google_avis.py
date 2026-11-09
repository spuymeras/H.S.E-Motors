"""Connexion à Google Business Profile : autorisation OAuth par agence et récupération
des avis. Aucune dépendance à Flask : tout ce qui touche à la base reçoit la connexion
sqlite en paramètre, et tous les appels réseau passent par appel_http (facile à
remplacer dans les tests, Google n'ayant pas de bac à sable)."""

import base64
import hashlib
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

from cryptography.fernet import Fernet, InvalidToken

SCOPES = "https://www.googleapis.com/auth/business.manage openid email"
URL_AUTORISATION = "https://accounts.google.com/o/oauth2/v2/auth"
URL_JETON = "https://oauth2.googleapis.com/token"
URL_REVOCATION = "https://oauth2.googleapis.com/revoke"
URL_PROFIL = "https://openidconnect.googleapis.com/v1/userinfo"
URL_COMPTES = "https://mybusinessaccountmanagement.googleapis.com/v1/accounts"
URL_LIEUX = "https://mybusinessbusinessinformation.googleapis.com/v1/{compte}/locations"
URL_AVIS = "https://mybusiness.googleapis.com/v4/{lieu}/reviews"

NOTES = {"ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5}
MAX_PAGES = 200
DELAI_SYNCHRO_HEURES = 24


class ErreurGoogle(Exception):
    """code : "reconnexion" (autorisation à refaire), "acces_api" (accès à l'API non
    accordé), "limite" (quota), "config" (réglage manquant) ou "autre"."""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


def est_configure():
    return bool(os.environ.get("GOOGLE_CLIENT_ID") and os.environ.get("GOOGLE_CLIENT_SECRET"))


def _identifiants():
    return os.environ.get("GOOGLE_CLIENT_ID", ""), os.environ.get("GOOGLE_CLIENT_SECRET", "")


def appel_http(methode, url, params=None, formulaire=None, jeton=None):
    if params:
        url = url + "?" + urllib.parse.urlencode(params)
    donnees = urllib.parse.urlencode(formulaire).encode() if formulaire is not None else None
    requete = urllib.request.Request(url, data=donnees, method=methode)
    if formulaire is not None:
        requete.add_header("Content-Type", "application/x-www-form-urlencoded")
    if jeton:
        requete.add_header("Authorization", f"Bearer {jeton}")
    try:
        with urllib.request.urlopen(requete, timeout=20) as reponse:
            corps = reponse.read().decode()
            return json.loads(corps) if corps else {}
    except urllib.error.HTTPError as e:
        corps = e.read().decode(errors="replace")
        try:
            detail = json.loads(corps)
        except ValueError:
            detail = {"raw": corps[:300]}
        raise erreur_depuis_reponse(e.code, detail)
    except urllib.error.URLError as e:
        raise ErreurGoogle("autre", f"Google injoignable : {e.reason}")


def erreur_depuis_reponse(statut, detail):
    erreur = detail.get("error") if isinstance(detail, dict) else None
    if erreur == "invalid_grant" or statut == 401:
        return ErreurGoogle("reconnexion", "Autorisation expirée ou révoquée : reconnectez le compte Google.")
    if statut == 429:
        return ErreurGoogle("limite", "Limite de requêtes Google atteinte, nouvel essai plus tard.")
    if statut == 403:
        return ErreurGoogle(
            "acces_api",
            "Accès refusé par Google : la demande d'accès à l'API Business Profile n'est peut-être "
            "pas encore approuvée, ou ce compte n'est pas propriétaire de la fiche.",
        )
    message = erreur.get("message") if isinstance(erreur, dict) else (erreur or "")
    return ErreurGoogle("autre", f"Erreur Google ({statut}) {message}".strip())


# ---------- Jeton d'actualisation chiffré au repos ----------

def _chiffreur(cle_secrete):
    cle = base64.urlsafe_b64encode(hashlib.sha256(cle_secrete.encode()).digest())
    return Fernet(cle)


def chiffrer(cle_secrete, texte):
    return _chiffreur(cle_secrete).encrypt(texte.encode()).decode()


def dechiffrer(cle_secrete, jeton):
    try:
        return _chiffreur(cle_secrete).decrypt(jeton.encode()).decode()
    except InvalidToken:
        raise ErreurGoogle("reconnexion", "Jeton illisible (clé secrète modifiée ?) : reconnectez le compte Google.")


# ---------- OAuth ----------

def url_autorisation(redirect_uri, etat):
    client_id, _ = _identifiants()
    return URL_AUTORISATION + "?" + urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",
        "state": etat,
    })


def echanger_code(code, redirect_uri):
    client_id, secret = _identifiants()
    return appel_http("POST", URL_JETON, formulaire={
        "code": code, "client_id": client_id, "client_secret": secret,
        "redirect_uri": redirect_uri, "grant_type": "authorization_code",
    })


def jeton_acces(refresh_token):
    client_id, secret = _identifiants()
    reponse = appel_http("POST", URL_JETON, formulaire={
        "refresh_token": refresh_token, "client_id": client_id,
        "client_secret": secret, "grant_type": "refresh_token",
    })
    return reponse["access_token"]


def email_du_compte(jeton):
    return appel_http("GET", URL_PROFIL, jeton=jeton).get("email", "")


def revoquer(refresh_token):
    try:
        appel_http("POST", URL_REVOCATION, formulaire={"token": refresh_token})
    except ErreurGoogle:
        pass


# ---------- Établissements et avis ----------

def _pages(url, jeton, cle, params_base):
    jetons_page = None
    for _ in range(MAX_PAGES):
        params = dict(params_base)
        if jetons_page:
            params["pageToken"] = jetons_page
        reponse = appel_http("GET", url, params=params, jeton=jeton)
        yield from reponse.get(cle, [])
        jetons_page = reponse.get("nextPageToken")
        if not jetons_page:
            return


def lister_lieux(jeton):
    lieux = []
    for compte in _pages(URL_COMPTES, jeton, "accounts", {"pageSize": 20}):
        url = URL_LIEUX.format(compte=compte["name"])
        for lieu in _pages(url, jeton, "locations", {"readMask": "name,title", "pageSize": 100}):
            lieux.append({
                "id": f'{compte["name"]}/{lieu["name"]}',
                "nom": lieu.get("title") or lieu["name"],
            })
    return lieux


def normaliser_avis(brut):
    auteur = brut.get("reviewer") or {}
    anonyme = auteur.get("isAnonymous") or not auteur.get("displayName")
    return {
        "review_id": brut["reviewId"],
        "auteur": "Anonyme" if anonyme else auteur["displayName"],
        "note": NOTES.get(brut.get("starRating"), 0),
        "commentaire": brut.get("comment", ""),
        "cree_le": brut.get("createTime"),
        "maj_le": brut.get("updateTime"),
    }


def lister_avis(jeton, lieu_id):
    url = URL_AVIS.format(lieu=lieu_id)
    return [normaliser_avis(a) for a in _pages(url, jeton, "reviews", {"pageSize": 50, "orderBy": "updateTime desc"})]


# ---------- Base de données ----------

def _maintenant():
    return datetime.utcnow().isoformat()


def enregistrer_connexion(db, cle_secrete, entreprise_id, refresh_token, email):
    db.execute(
        """INSERT INTO google_connexions (entreprise_id, email, refresh_token, connecte_le)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(entreprise_id) DO UPDATE SET
             email = excluded.email, refresh_token = excluded.refresh_token, connecte_le = excluded.connecte_le,
             lieu_id = NULL, lieu_nom = NULL, lieux_json = NULL, derniere_erreur = NULL""",
        (entreprise_id, email, chiffrer(cle_secrete, refresh_token), _maintenant()),
    )
    db.commit()


def charger_lieux(db, cle_secrete, entreprise_id):
    """Interroge Google pour la liste des établissements du compte connecté. Si un seul
    existe, il est choisi d'office. Retourne la liste ; lève ErreurGoogle en cas d'échec
    (le message est aussi mémorisé pour l'affichage)."""
    connexion = db.execute("SELECT * FROM google_connexions WHERE entreprise_id = ?", (entreprise_id,)).fetchone()
    if connexion is None:
        raise ErreurGoogle("autre", "Aucun compte Google connecté pour cette entreprise.")
    try:
        jeton = jeton_acces(dechiffrer(cle_secrete, connexion["refresh_token"]))
        lieux = lister_lieux(jeton)
    except ErreurGoogle as e:
        db.execute("UPDATE google_connexions SET derniere_erreur = ? WHERE entreprise_id = ?", (e.message, entreprise_id))
        db.commit()
        raise
    lieu_id, lieu_nom = (lieux[0]["id"], lieux[0]["nom"]) if len(lieux) == 1 else (None, None)
    db.execute(
        "UPDATE google_connexions SET lieux_json = ?, lieu_id = ?, lieu_nom = ?, derniere_erreur = NULL WHERE entreprise_id = ?",
        (json.dumps(lieux), lieu_id, lieu_nom, entreprise_id),
    )
    db.commit()
    return lieux


def synchroniser(db, cle_secrete, entreprise_id):
    """Récupère tous les avis de la fiche et les enregistre (mise à jour si déjà connus).
    Retourne {"total": n, "nouveaux": n}."""
    connexion = db.execute("SELECT * FROM google_connexions WHERE entreprise_id = ?", (entreprise_id,)).fetchone()
    if connexion is None or not connexion["lieu_id"]:
        raise ErreurGoogle("autre", "Aucun établissement choisi pour cette entreprise.")
    try:
        jeton = jeton_acces(dechiffrer(cle_secrete, connexion["refresh_token"]))
        avis = lister_avis(jeton, connexion["lieu_id"])
    except ErreurGoogle as e:
        # Nouvel essai automatique dans ~1 h plutôt que dans 24 h.
        reessai = (datetime.utcnow() - timedelta(hours=DELAI_SYNCHRO_HEURES - 1)).isoformat()
        db.execute(
            "UPDATE google_connexions SET derniere_erreur = ?, derniere_tentative = ? WHERE entreprise_id = ?",
            (e.message, reessai, entreprise_id),
        )
        db.commit()
        raise

    maintenant = _maintenant()
    nouveaux = 0
    for a in avis:
        existe = db.execute("SELECT 1 FROM avis_google_recus WHERE review_id = ?", (a["review_id"],)).fetchone()
        if existe is None:
            nouveaux += 1
        db.execute(
            """INSERT INTO avis_google_recus (review_id, entreprise_id, auteur, note, commentaire, cree_le, maj_le, recupere_le)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(review_id) DO UPDATE SET
                 auteur = excluded.auteur, note = excluded.note, commentaire = excluded.commentaire, maj_le = excluded.maj_le""",
            (a["review_id"], entreprise_id, a["auteur"], a["note"], a["commentaire"], a["cree_le"], a["maj_le"], maintenant),
        )
    db.execute(
        "UPDATE google_connexions SET derniere_sync = ?, derniere_tentative = ?, derniere_erreur = NULL WHERE entreprise_id = ?",
        (maintenant, maintenant, entreprise_id),
    )
    db.commit()
    return {"total": len(avis), "nouveaux": nouveaux}


def reclamer_synchro(db, entreprise_id):
    """Réservation atomique : avec plusieurs workers gunicorn, un seul obtient la
    synchronisation quotidienne d'une entreprise (le premier à poser la tentative)."""
    limite = (datetime.utcnow() - timedelta(hours=DELAI_SYNCHRO_HEURES)).isoformat()
    curseur = db.execute(
        """UPDATE google_connexions SET derniere_tentative = ?
           WHERE entreprise_id = ? AND lieu_id IS NOT NULL
             AND (derniere_tentative IS NULL OR derniere_tentative < ?)""",
        (_maintenant(), entreprise_id, limite),
    )
    db.commit()
    return curseur.rowcount == 1


def deconnecter(db, cle_secrete, entreprise_id):
    connexion = db.execute("SELECT refresh_token FROM google_connexions WHERE entreprise_id = ?", (entreprise_id,)).fetchone()
    if connexion is None:
        return
    try:
        revoquer(dechiffrer(cle_secrete, connexion["refresh_token"]))
    except ErreurGoogle:
        pass
    db.execute("DELETE FROM google_connexions WHERE entreprise_id = ?", (entreprise_id,))
    db.commit()
