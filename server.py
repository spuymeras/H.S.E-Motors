import json
import os
import smtplib
import sqlite3
import secrets
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from functools import wraps
from pathlib import Path

from flask import Flask, g, jsonify, request, send_from_directory, session
from werkzeug.security import check_password_hash, generate_password_hash

try:
    import email_config
except ModuleNotFoundError:
    # Absent en production par design (voir .gitignore) : les réglages viennent
    # alors uniquement des variables d'environnement, via parametre_email().
    email_config = None

BASE_DIR = Path(__file__).resolve().parent

# En production, DATA_DIR pointe vers un volume persistant (ex: /data sur Railway) pour que
# la base de données survive aux redéploiements. En local, elle reste à côté du code.
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "app.db"
SECRET_KEY_PATH = BASE_DIR / "secret.key"

app = Flask(__name__, static_folder="static", static_url_path="")

# En production, définissez la variable d'environnement SECRET_KEY (l'hébergeur la fournit
# dans son tableau de bord). En local, une clé est générée une fois et stockée dans secret.key.
if os.environ.get("SECRET_KEY"):
    app.secret_key = os.environ["SECRET_KEY"]
else:
    if not SECRET_KEY_PATH.exists():
        SECRET_KEY_PATH.write_text(secrets.token_hex(32))
    app.secret_key = SECRET_KEY_PATH.read_text().strip()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")


def parametre_email(nom):
    """Variable d'environnement en priorité (production), sinon email_config.py (local)."""
    valeur_locale = getattr(email_config, nom, "") if email_config else ""
    return os.environ.get(nom) or valeur_locale


# ---------- Base de données ----------

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")
        g.db.execute("PRAGMA busy_timeout = 5000")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


SCHEMA = """
CREATE TABLE IF NOT EXISTS entreprises (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    couleur TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS commerciaux (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    entreprise_id INTEGER NOT NULL REFERENCES entreprises(id),
    taux REAL,
    actif INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS utilisateurs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    identifiant TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('admin', 'commercial', 'animateur')),
    commercial_id INTEGER REFERENCES commerciaux(id),
    nom TEXT,
    kpi_order TEXT
);

CREATE TABLE IF NOT EXISTS dossiers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    commercial_id INTEGER NOT NULL REFERENCES commerciaux(id),
    date TEXT NOT NULL,
    client TEXT NOT NULL,
    voiture TEXT NOT NULL,
    plaque TEXT,
    garantie_achat REAL NOT NULL DEFAULT 0,
    garantie_prix_vendu REAL NOT NULL DEFAULT 0,
    mandat_total REAL NOT NULL DEFAULT 0,
    frais_intermediation REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS reinitialisations_mdp (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    utilisateur_id INTEGER NOT NULL REFERENCES utilisateurs(id),
    code_hash TEXT NOT NULL,
    expire_le TEXT NOT NULL,
    utilise INTEGER NOT NULL DEFAULT 0
);
"""


def migrer_db(db):
    colonnes = {row["name"] for row in db.execute("PRAGMA table_info(dossiers)")}
    if "frais_intermediation" not in colonnes:
        db.execute("ALTER TABLE dossiers ADD COLUMN frais_intermediation REAL NOT NULL DEFAULT 0")

    colonnes_utilisateurs = {row["name"] for row in db.execute("PRAGMA table_info(utilisateurs)")}
    if "nom" not in colonnes_utilisateurs:
        db.execute("ALTER TABLE utilisateurs ADD COLUMN nom TEXT")
    if "kpi_order" not in colonnes_utilisateurs:
        db.execute("ALTER TABLE utilisateurs ADD COLUMN kpi_order TEXT")

    schema_utilisateurs = db.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'utilisateurs'"
    ).fetchone()[0]
    if "'animateur'" not in schema_utilisateurs:
        db.executescript(
            """
            CREATE TABLE utilisateurs_nouveau (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                identifiant TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'commercial', 'animateur')),
                commercial_id INTEGER REFERENCES commerciaux(id),
                nom TEXT,
                kpi_order TEXT
            );
            INSERT INTO utilisateurs_nouveau (id, identifiant, password_hash, role, commercial_id, nom, kpi_order)
                SELECT id, identifiant, password_hash, role, commercial_id, nom, kpi_order FROM utilisateurs;
            DROP TABLE utilisateurs;
            ALTER TABLE utilisateurs_nouveau RENAME TO utilisateurs;
            """
        )

    schema_commerciaux = db.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'commerciaux'"
    ).fetchone()[0]
    if "taux REAL NOT NULL" in schema_commerciaux:
        db.executescript(
            """
            CREATE TABLE commerciaux_nouveau (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT NOT NULL,
                entreprise_id INTEGER NOT NULL REFERENCES entreprises(id),
                taux REAL,
                actif INTEGER NOT NULL DEFAULT 1
            );
            INSERT INTO commerciaux_nouveau (id, nom, entreprise_id, taux, actif)
                SELECT id, nom, entreprise_id, taux, actif FROM commerciaux;
            DROP TABLE commerciaux;
            ALTER TABLE commerciaux_nouveau RENAME TO commerciaux;
            """
        )


def init_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    migrer_db(db)
    already_seeded = db.execute("SELECT COUNT(*) FROM entreprises").fetchone()[0] > 0
    if not already_seeded:
        seed(db)
    db.commit()
    db.close()


def seed(db):
    entreprises = [
        ("Norda Bâtiment", "#E8A33D"),
        ("Lumis Énergie", "#4FD1C5"),
        ("Verrel Distribution", "#C58AF0"),
    ]
    db.executemany("INSERT INTO entreprises (nom, couleur) VALUES (?, ?)", entreprises)
    ent_ids = [r[0] for r in db.execute("SELECT id FROM entreprises ORDER BY id")]

    commerciaux = [
        ("Awa Diallo", ent_ids[0], 0.08, "awa.diallo"),
        ("Marc Lefèvre", ent_ids[0], 0.07, "marc.lefevre"),
        ("Julie Rambert", ent_ids[1], 0.10, "julie.rambert"),
        ("Karim Belhadj", ent_ids[2], 0.06, "karim.belhadj"),
    ]
    commercial_ids = {}
    for nom, ent_id, taux, identifiant in commerciaux:
        cur = db.execute(
            "INSERT INTO commerciaux (nom, entreprise_id, taux) VALUES (?, ?, ?)",
            (nom, ent_id, taux),
        )
        commercial_ids[identifiant] = cur.lastrowid
        db.execute(
            "INSERT INTO utilisateurs (identifiant, password_hash, role, commercial_id) VALUES (?, ?, 'commercial', ?)",
            (identifiant, generate_password_hash("commercial123", method="pbkdf2:sha256"), cur.lastrowid),
        )

    db.execute(
        "INSERT INTO utilisateurs (identifiant, password_hash, role, commercial_id) VALUES (?, ?, 'admin', NULL)",
        ("admin", generate_password_hash("admin123", method="pbkdf2:sha256")),
    )

    ventes = [
        (commercial_ids["awa.diallo"], "2026-07-03", "M. Beaumont", "Peugeot 3008", "EF-482-GT", 15200, 18400, 1200),
        (commercial_ids["awa.diallo"], "2026-07-14", "Ateliers du Nord", "Renault Trafic", "AB-113-CD", 8100, 9200, 650),
        (commercial_ids["marc.lefevre"], "2026-07-08", "Mme Ferran", "BMW Série 3", "GH-905-JK", 23800, 26500, 1800),
        (commercial_ids["julie.rambert"], "2026-07-11", "EcoVolt Sarl", "Tesla Model 3", "LM-227-NO", 37500, 41200, 2400),
        (commercial_ids["julie.rambert"], "2026-07-22", "Résidence Les Tilleuls", "Citroën C3", "PQ-664-RS", 13900, 15800, 900),
        (commercial_ids["karim.belhadj"], "2026-07-18", "Distri Ouest", "Ford Transit", "TU-338-VW", 10800, 12300, 780),
    ]
    db.executemany(
        """INSERT INTO dossiers
           (commercial_id, date, client, voiture, plaque, garantie_achat, garantie_prix_vendu, mandat_total)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        ventes,
    )


# ---------- Auth helpers ----------

def current_user():
    uid = session.get("uid")
    if not uid:
        return None
    db = get_db()
    return db.execute("SELECT * FROM utilisateurs WHERE id = ?", (uid,)).fetchone()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if user is None:
            return jsonify(error="Non authentifié"), 401
        g.user = user
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if user is None:
            return jsonify(error="Non authentifié"), 401
        if user["role"] != "admin":
            return jsonify(error="Réservé aux administrateurs"), 403
        g.user = user
        return fn(*args, **kwargs)
    return wrapper


def ecriture_requise(fn):
    """Comme login_required, mais bloque le rôle animateur (accès lecture seule)."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if user is None:
            return jsonify(error="Non authentifié"), 401
        if user["role"] == "animateur":
            return jsonify(error="Accès en lecture seule"), 403
        g.user = user
        return fn(*args, **kwargs)
    return wrapper


def peut_tout_voir(user):
    return user["role"] in ("admin", "animateur")


def user_public(user, db):
    if user["role"] == "admin":
        kpi_order = json.loads(user["kpi_order"]) if user["kpi_order"] else None
        return {
            "id": user["id"],
            "role": "admin",
            "nom": user["nom"] or "Administrateur",
            "identifiant": user["identifiant"],
            "kpi_order": kpi_order,
        }
    if user["role"] == "animateur":
        return {"id": user["id"], "role": "animateur", "nom": user["nom"] or "Animateur", "identifiant": user["identifiant"]}
    com = db.execute("SELECT * FROM commerciaux WHERE id = ?", (user["commercial_id"],)).fetchone()
    return {
        "id": user["id"],
        "role": "commercial",
        "nom": com["nom"],
        "identifiant": user["identifiant"],
        "commercial_id": com["id"],
        "entreprise_id": com["entreprise_id"],
        "taux": com["taux"],
    }


# ---------- Routes: auth ----------

@app.post("/api/login")
def login():
    data = request.get_json(silent=True) or {}
    identifiant = (data.get("identifiant") or "").strip()
    mot_de_passe = data.get("mot_de_passe") or ""
    db = get_db()
    user = db.execute("SELECT * FROM utilisateurs WHERE identifiant = ?", (identifiant,)).fetchone()
    if user is None or not check_password_hash(user["password_hash"], mot_de_passe):
        return jsonify(error="Identifiant ou mot de passe incorrect"), 401
    if user["role"] == "commercial":
        com = db.execute("SELECT actif FROM commerciaux WHERE id = ?", (user["commercial_id"],)).fetchone()
        if com is None or not com["actif"]:
            return jsonify(error="Ce compte est désactivé"), 403
    session.clear()
    session["uid"] = user["id"]
    session.permanent = True
    return jsonify(user_public(user, db))


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify(ok=True)


def envoyer_code_recuperation(code):
    smtp_user = parametre_email("SMTP_USER")
    smtp_password = parametre_email("SMTP_PASSWORD")
    admin_recovery_email = parametre_email("ADMIN_RECOVERY_EMAIL")
    smtp_host = parametre_email("SMTP_HOST") or "smtp.gmail.com"
    smtp_port = int(parametre_email("SMTP_PORT") or 587)

    if not smtp_user or not smtp_password or not admin_recovery_email:
        raise RuntimeError(
            "L'envoi d'email n'est pas configuré. Remplissez email_config.py (SMTP_USER, "
            "SMTP_PASSWORD, ADMIN_RECOVERY_EMAIL) puis relancez le serveur."
        )
    message = MIMEText(
        f"Voici votre code de récupération pour H.S.E Motors : {code}\n\n"
        "Ce code est valable 15 minutes. Si vous n'êtes pas à l'origine de cette demande, ignorez cet email."
    )
    message["Subject"] = "H.S.E Motors — Code de récupération du mot de passe admin"
    message["From"] = smtp_user
    message["To"] = admin_recovery_email

    with smtplib.SMTP(smtp_host, smtp_port) as smtp:
        smtp.starttls()
        smtp.login(smtp_user, smtp_password)
        smtp.send_message(message)


@app.post("/api/mot-de-passe-oublie")
def demander_code_recuperation():
    data = request.get_json(silent=True) or {}
    identifiant = (data.get("identifiant") or "").strip()
    db = get_db()
    user = db.execute(
        "SELECT * FROM utilisateurs WHERE identifiant = ? AND role = 'admin'", (identifiant,)
    ).fetchone()

    # Réponse identique que le compte existe ou non, pour ne pas révéler les identifiants valides.
    reponse_generique = jsonify(
        ok=True, message="Si ce compte existe, un code de récupération a été envoyé par email."
    )

    if user is None:
        return reponse_generique

    code = f"{secrets.randbelow(1_000_000):06d}"
    expire_le = (datetime.utcnow() + timedelta(minutes=15)).isoformat()
    db.execute(
        "INSERT INTO reinitialisations_mdp (utilisateur_id, code_hash, expire_le) VALUES (?, ?, ?)",
        (user["id"], generate_password_hash(code, method="pbkdf2:sha256"), expire_le),
    )
    db.commit()

    try:
        envoyer_code_recuperation(code)
    except Exception as e:
        return jsonify(error=str(e)), 500

    return reponse_generique


@app.post("/api/reinitialiser-mot-de-passe")
def reinitialiser_mot_de_passe():
    data = request.get_json(silent=True) or {}
    identifiant = (data.get("identifiant") or "").strip()
    code = (data.get("code") or "").strip()
    nouveau_mot_de_passe = (data.get("nouveau_mot_de_passe") or "").strip()

    if not nouveau_mot_de_passe:
        return jsonify(error="Le nouveau mot de passe est requis"), 400

    db = get_db()
    user = db.execute(
        "SELECT * FROM utilisateurs WHERE identifiant = ? AND role = 'admin'", (identifiant,)
    ).fetchone()
    if user is None:
        return jsonify(error="Code invalide ou expiré"), 400

    maintenant = datetime.utcnow().isoformat()
    demandes = db.execute(
        """SELECT * FROM reinitialisations_mdp
           WHERE utilisateur_id = ? AND utilise = 0 AND expire_le > ?
           ORDER BY id DESC""",
        (user["id"], maintenant),
    ).fetchall()

    demande_valide = next((d for d in demandes if check_password_hash(d["code_hash"], code)), None)
    if demande_valide is None:
        return jsonify(error="Code invalide ou expiré"), 400

    db.execute(
        "UPDATE utilisateurs SET password_hash = ? WHERE id = ?",
        (generate_password_hash(nouveau_mot_de_passe, method="pbkdf2:sha256"), user["id"]),
    )
    db.execute("UPDATE reinitialisations_mdp SET utilise = 1 WHERE id = ?", (demande_valide["id"],))
    db.commit()
    return jsonify(ok=True)


@app.get("/api/me")
@login_required
def me():
    return jsonify(user_public(g.user, get_db()))


@app.put("/api/preferences/kpi-order")
@admin_required
def sauvegarder_ordre_kpi():
    data = request.get_json(silent=True) or {}
    ordre = data.get("order")
    if not isinstance(ordre, list) or not all(isinstance(x, str) for x in ordre):
        return jsonify(error="Format d'ordre invalide"), 400

    db = get_db()
    db.execute("UPDATE utilisateurs SET kpi_order = ? WHERE id = ?", (json.dumps(ordre), g.user["id"]))
    db.commit()
    return jsonify(ok=True, order=ordre)


# ---------- Routes: entreprises ----------

@app.get("/api/entreprises")
@login_required
def list_entreprises():
    db = get_db()
    rows = db.execute("SELECT * FROM entreprises ORDER BY nom").fetchall()
    return jsonify([dict(r) for r in rows])


@app.post("/api/entreprises")
@admin_required
def create_entreprise():
    data = request.get_json(silent=True) or {}
    nom = (data.get("nom") or "").strip()
    couleur = (data.get("couleur") or "#8B93A1").strip()
    if not nom:
        return jsonify(error="Le nom est requis"), 400
    db = get_db()
    count = db.execute("SELECT COUNT(*) FROM entreprises").fetchone()[0]
    if count >= 5:
        return jsonify(error="Limite de 5 entreprises atteinte"), 400
    cur = db.execute("INSERT INTO entreprises (nom, couleur) VALUES (?, ?)", (nom, couleur))
    db.commit()
    return jsonify(id=cur.lastrowid, nom=nom, couleur=couleur), 201


@app.put("/api/entreprises/<int:entreprise_id>")
@admin_required
def update_entreprise(entreprise_id):
    db = get_db()
    row = db.execute("SELECT * FROM entreprises WHERE id = ?", (entreprise_id,)).fetchone()
    if row is None:
        return jsonify(error="Entreprise introuvable"), 404

    data = request.get_json(silent=True) or {}
    nom = (data.get("nom") or row["nom"]).strip()
    couleur = (data.get("couleur") or row["couleur"]).strip()
    if not nom:
        return jsonify(error="Le nom est requis"), 400

    db.execute("UPDATE entreprises SET nom = ?, couleur = ? WHERE id = ?", (nom, couleur, entreprise_id))
    db.commit()
    return jsonify(id=entreprise_id, nom=nom, couleur=couleur)


# ---------- Routes: commerciaux ----------

SEUIL_TAUX_AUTO = 8000.0
TAUX_AUTO_BAS = 0.20
TAUX_AUTO_HAUT = 0.35


def total_ht_ligne(row):
    tca = max(0, row["garantie_prix_vendu"] - row["garantie_achat"]) * 0.18
    return (row["mandat_total"] - row["garantie_achat"] - tca - CASH_SENTINEL) / 1.2


def ca_ht_net_mensuel(db, commercial_id, mois):
    rows = db.execute(
        "SELECT garantie_achat, garantie_prix_vendu, mandat_total FROM dossiers "
        "WHERE commercial_id = ? AND substr(date, 1, 7) = ?",
        (commercial_id, mois),
    ).fetchall()
    return sum(total_ht_ligne(r) for r in rows)


def taux_automatique(db, commercial_id, mois):
    ca = ca_ht_net_mensuel(db, commercial_id, mois)
    return TAUX_AUTO_HAUT if ca > SEUIL_TAUX_AUTO else TAUX_AUTO_BAS


def taux_effectif(db, com, mois):
    if com["taux"] is not None:
        return com["taux"]
    return taux_automatique(db, com["id"], mois)


def commercial_dict(row, db):
    mois_courant = datetime.utcnow().strftime("%Y-%m")
    return {
        "id": row["id"],
        "nom": row["nom"],
        "entreprise_id": row["entreprise_id"],
        "taux": row["taux"],
        "taux_auto": row["taux"] is None,
        "taux_effectif": taux_effectif(db, row, mois_courant),
        "actif": bool(row["actif"]),
    }


@app.get("/api/commerciaux")
@login_required
def list_commerciaux():
    db = get_db()
    if peut_tout_voir(g.user):
        rows = db.execute("SELECT * FROM commerciaux ORDER BY nom").fetchall()
    else:
        rows = db.execute("SELECT * FROM commerciaux WHERE id = ?", (g.user["commercial_id"],)).fetchall()
    return jsonify([commercial_dict(r, db) for r in rows])


def parser_taux_optionnel(data):
    """Retourne (taux, erreur). taux est None si absent/vide (mode automatique)."""
    if "taux" not in data or data["taux"] in (None, ""):
        return None, None
    try:
        return float(data["taux"]), None
    except (TypeError, ValueError):
        return None, "Taux invalide"


@app.post("/api/commerciaux")
@admin_required
def create_commercial():
    data = request.get_json(silent=True) or {}
    nom = (data.get("nom") or "").strip()
    entreprise_id = data.get("entreprise_id")
    identifiant = (data.get("identifiant") or "").strip()
    mot_de_passe = data.get("mot_de_passe") or ""

    if not nom or not entreprise_id or not identifiant or not mot_de_passe:
        return jsonify(error="Nom, entreprise, identifiant et mot de passe sont requis"), 400

    taux, erreur = parser_taux_optionnel(data)
    if erreur:
        return jsonify(error=erreur), 400

    db = get_db()
    if db.execute("SELECT 1 FROM utilisateurs WHERE identifiant = ?", (identifiant,)).fetchone():
        return jsonify(error="Cet identifiant existe déjà"), 400
    if not db.execute("SELECT 1 FROM entreprises WHERE id = ?", (entreprise_id,)).fetchone():
        return jsonify(error="Entreprise inconnue"), 400

    cur = db.execute(
        "INSERT INTO commerciaux (nom, entreprise_id, taux) VALUES (?, ?, ?)",
        (nom, entreprise_id, taux),
    )
    commercial_id = cur.lastrowid
    db.execute(
        "INSERT INTO utilisateurs (identifiant, password_hash, role, commercial_id) VALUES (?, ?, 'commercial', ?)",
        (identifiant, generate_password_hash(mot_de_passe, method="pbkdf2:sha256"), commercial_id),
    )
    db.commit()
    row = db.execute("SELECT * FROM commerciaux WHERE id = ?", (commercial_id,)).fetchone()
    return jsonify(commercial_dict(row, db)), 201


@app.put("/api/commerciaux/<int:commercial_id>")
@admin_required
def update_commercial(commercial_id):
    data = request.get_json(silent=True) or {}
    db = get_db()
    row = db.execute("SELECT * FROM commerciaux WHERE id = ?", (commercial_id,)).fetchone()
    if row is None:
        return jsonify(error="Commercial introuvable"), 404

    nom = data.get("nom", row["nom"])
    actif = data.get("actif", bool(row["actif"]))
    entreprise_id = data.get("entreprise_id", row["entreprise_id"])

    if "taux" in data:
        taux, erreur = parser_taux_optionnel(data)
        if erreur:
            return jsonify(error=erreur), 400
    else:
        taux = row["taux"]

    if not db.execute("SELECT 1 FROM entreprises WHERE id = ?", (entreprise_id,)).fetchone():
        return jsonify(error="Entreprise inconnue"), 400

    nouveau_mot_de_passe = (data.get("mot_de_passe") or "").strip()

    db.execute(
        "UPDATE commerciaux SET nom = ?, taux = ?, actif = ?, entreprise_id = ? WHERE id = ?",
        (nom, taux, 1 if actif else 0, entreprise_id, commercial_id),
    )
    if nouveau_mot_de_passe:
        db.execute(
            "UPDATE utilisateurs SET password_hash = ? WHERE commercial_id = ?",
            (generate_password_hash(nouveau_mot_de_passe, method="pbkdf2:sha256"), commercial_id),
        )
    db.commit()
    row = db.execute("SELECT * FROM commerciaux WHERE id = ?", (commercial_id,)).fetchone()
    return jsonify(commercial_dict(row, db))


# ---------- Routes: animateurs (accès lecture seule, toutes entreprises) ----------

def animateur_dict(row):
    return {"id": row["id"], "identifiant": row["identifiant"], "nom": row["nom"]}


@app.get("/api/animateurs")
@admin_required
def list_animateurs():
    db = get_db()
    rows = db.execute("SELECT * FROM utilisateurs WHERE role = 'animateur' ORDER BY nom").fetchall()
    return jsonify([animateur_dict(r) for r in rows])


@app.post("/api/animateurs")
@admin_required
def create_animateur():
    data = request.get_json(silent=True) or {}
    nom = (data.get("nom") or "").strip()
    identifiant = (data.get("identifiant") or "").strip()
    mot_de_passe = data.get("mot_de_passe") or ""

    if not nom or not identifiant or not mot_de_passe:
        return jsonify(error="Nom, identifiant et mot de passe sont requis"), 400

    db = get_db()
    if db.execute("SELECT 1 FROM utilisateurs WHERE identifiant = ?", (identifiant,)).fetchone():
        return jsonify(error="Cet identifiant existe déjà"), 400

    cur = db.execute(
        "INSERT INTO utilisateurs (identifiant, password_hash, role, nom) VALUES (?, ?, 'animateur', ?)",
        (identifiant, generate_password_hash(mot_de_passe, method="pbkdf2:sha256"), nom),
    )
    db.commit()
    row = db.execute("SELECT * FROM utilisateurs WHERE id = ?", (cur.lastrowid,)).fetchone()
    return jsonify(animateur_dict(row)), 201


# ---------- Routes: dossiers ----------

CASH_SENTINEL = 43.20


def dossier_dict(row, com, ent, db):
    tca = max(0, row["garantie_prix_vendu"] - row["garantie_achat"]) * 0.18
    total_ht = (row["mandat_total"] - row["garantie_achat"] - tca - CASH_SENTINEL) / 1.2
    taux = taux_effectif(db, com, row["date"][:7])
    commission = total_ht * taux
    return {
        "id": row["id"],
        "commercial_id": row["commercial_id"],
        "commercial_nom": com["nom"],
        "entreprise_id": ent["id"],
        "entreprise_nom": ent["nom"],
        "entreprise_couleur": ent["couleur"],
        "date": row["date"],
        "client": row["client"],
        "voiture": row["voiture"],
        "plaque": row["plaque"],
        "garantie_achat": row["garantie_achat"],
        "garantie_prix_vendu": row["garantie_prix_vendu"],
        "mandat_total": row["mandat_total"],
        "frais_intermediation": row["frais_intermediation"],
        "tca": tca,
        "cash_sentinel": CASH_SENTINEL,
        "total_ht": total_ht,
        "commission": commission,
    }


def enrich(db, rows):
    result = []
    for row in rows:
        com = db.execute("SELECT * FROM commerciaux WHERE id = ?", (row["commercial_id"],)).fetchone()
        ent = db.execute("SELECT * FROM entreprises WHERE id = ?", (com["entreprise_id"],)).fetchone()
        result.append(dossier_dict(row, com, ent, db))
    return result


@app.get("/api/dossiers")
@login_required
def list_dossiers():
    db = get_db()
    if peut_tout_voir(g.user):
        entreprise_id = request.args.get("entreprise_id")
        if entreprise_id and entreprise_id != "toutes":
            rows = db.execute(
                """SELECT d.* FROM dossiers d
                   JOIN commerciaux c ON c.id = d.commercial_id
                   WHERE c.entreprise_id = ? ORDER BY d.date DESC""",
                (entreprise_id,),
            ).fetchall()
        else:
            rows = db.execute("SELECT * FROM dossiers ORDER BY date DESC").fetchall()
    else:
        rows = db.execute(
            "SELECT * FROM dossiers WHERE commercial_id = ? ORDER BY date DESC",
            (g.user["commercial_id"],),
        ).fetchall()
    return jsonify(enrich(db, rows))


def resolve_commercial_id(data):
    if g.user["role"] == "admin":
        commercial_id = data.get("commercial_id")
        if not commercial_id:
            return None, (jsonify(error="commercial_id requis"), 400)
        return commercial_id, None
    return g.user["commercial_id"], None


@app.post("/api/dossiers")
@ecriture_requise
def create_dossier():
    data = request.get_json(silent=True) or {}
    commercial_id, err = resolve_commercial_id(data)
    if err:
        return err

    client = (data.get("client") or "").strip()
    voiture = (data.get("voiture") or "").strip()
    mandat_total = data.get("mandat_total")
    if not client or not voiture or mandat_total is None or float(mandat_total) <= 0:
        return jsonify(error="Client, voiture et mandat total (>0) sont requis"), 400

    db = get_db()
    if not db.execute("SELECT 1 FROM commerciaux WHERE id = ?", (commercial_id,)).fetchone():
        return jsonify(error="Commercial inconnu"), 400

    cur = db.execute(
        """INSERT INTO dossiers
           (commercial_id, date, client, voiture, plaque, garantie_achat, garantie_prix_vendu, mandat_total, frais_intermediation)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            commercial_id,
            data.get("date") or str(date.today()),
            client,
            voiture,
            (data.get("plaque") or "").strip().upper(),
            float(data.get("garantie_achat") or 0),
            float(data.get("garantie_prix_vendu") or 0),
            float(mandat_total),
            float(data.get("frais_intermediation") or 0),
        ),
    )
    db.commit()
    row = db.execute("SELECT * FROM dossiers WHERE id = ?", (cur.lastrowid,)).fetchone()
    return jsonify(enrich(db, [row])[0]), 201


def load_dossier_for_user(db, dossier_id):
    row = db.execute("SELECT * FROM dossiers WHERE id = ?", (dossier_id,)).fetchone()
    if row is None:
        return None
    if g.user["role"] != "admin" and row["commercial_id"] != g.user["commercial_id"]:
        return None
    return row


@app.put("/api/dossiers/<int:dossier_id>")
@ecriture_requise
def update_dossier(dossier_id):
    db = get_db()
    row = load_dossier_for_user(db, dossier_id)
    if row is None:
        return jsonify(error="Dossier introuvable"), 404

    data = request.get_json(silent=True) or {}
    commercial_id = row["commercial_id"]
    if g.user["role"] == "admin" and data.get("commercial_id"):
        commercial_id = data["commercial_id"]

    db.execute(
        """UPDATE dossiers SET commercial_id = ?, date = ?, client = ?, voiture = ?, plaque = ?,
           garantie_achat = ?, garantie_prix_vendu = ?, mandat_total = ?, frais_intermediation = ? WHERE id = ?""",
        (
            commercial_id,
            data.get("date", row["date"]),
            (data.get("client", row["client"]) or "").strip(),
            (data.get("voiture", row["voiture"]) or "").strip(),
            (data.get("plaque", row["plaque"]) or "").strip().upper(),
            float(data.get("garantie_achat", row["garantie_achat"]) or 0),
            float(data.get("garantie_prix_vendu", row["garantie_prix_vendu"]) or 0),
            float(data.get("mandat_total", row["mandat_total"]) or 0),
            float(data.get("frais_intermediation", row["frais_intermediation"]) or 0),
            dossier_id,
        ),
    )
    db.commit()
    row = db.execute("SELECT * FROM dossiers WHERE id = ?", (dossier_id,)).fetchone()
    return jsonify(enrich(db, [row])[0])


@app.delete("/api/dossiers/<int:dossier_id>")
@ecriture_requise
def delete_dossier(dossier_id):
    db = get_db()
    row = load_dossier_for_user(db, dossier_id)
    if row is None:
        return jsonify(error="Dossier introuvable"), 404
    db.execute("DELETE FROM dossiers WHERE id = ?", (dossier_id,))
    db.commit()
    return jsonify(ok=True)


# ---------- Frontend ----------

@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


# Exécuté à l'import du module : garantit que la base est prête à la fois en local
# (python3 server.py) et en production (gunicorn server:app, qui n'exécute jamais __main__).
init_db()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8731, debug=True)
