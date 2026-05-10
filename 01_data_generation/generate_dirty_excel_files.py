import os
import random
from datetime import timedelta

import numpy as np
import pandas as pd


# =====================================================
# PARAMÈTRES GÉNÉRAUX
# =====================================================

SEED = 42
N_ROWS = 20_000

random.seed(SEED)
np.random.seed(SEED)

OUTPUT_DIR = "suez_normandie_dirty_files"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# =====================================================
# LISTES MÉTIER
# =====================================================

DEPARTEMENTS = ["Calvados", "Seine-Maritime", "Eure", "Manche", "Orne"]

VILLES = [
    "Rouen", "Caen", "Le Havre", "Évreux", "Cherbourg",
    "Alençon", "Dieppe", "Lisieux", "Saint-Lô", "Bayeux"
]

SITES = [
    "UVE Caen Colombelles",
    "Centre de tri Rouen",
    "Plateforme déchets Le Havre",
    "Centre de valorisation Évreux",
    "Site traitement Alençon",
    "Centre recyclage Dieppe"
]

TYPES_CLIENT = [
    "Collectivité",
    "Industrie",
    "Entreprise privée",
    "Syndicat mixte",
    "Communauté de communes",
    "Administration"
]

TYPES_CONTRAT = [
    "Collecte déchets",
    "Traitement déchets",
    "Valorisation énergétique",
    "Maintenance site",
    "Prestation industrielle",
    "Recyclage"
]

TYPES_DECHET = [
    "Ordures ménagères",
    "Déchets industriels banals",
    "Déchets verts",
    "Encombrants",
    "CSR",
    "Plastiques",
    "Métaux",
    "Bois",
    "Papier carton",
    "Déchets dangereux"
]

STATUTS_CLIENT = ["Actif", "Inactif", "Suspendu", "Résilié"]

STATUTS_CONTRAT = [
    "Actif",
    "Expiré",
    "Résilié",
    "Suspendu",
    "En renouvellement"
]

STATUTS_FACTURE = [
    "Payée",
    "En attente",
    "En retard",
    "Partiellement payée",
    "Litige"
]

STATUTS_INTERVENTION = [
    "Planifiée",
    "En cours",
    "Clôturée",
    "Annulée",
    "En retard"
]

STATUTS_PRODUCTION = [
    "Normale",
    "Dégradée",
    "Arrêt technique",
    "Maintenance",
    "Incident"
]

RESPONSABLES = [
    "A. Martin", "S. Bernard", "M. Dubois", "L. Petit",
    "K. Moreau", "N. Laurent", "F. Leroy", "C. Simon"
]

SERVICES = [
    "Finance",
    "Exploitation",
    "Maintenance",
    "Commerce",
    "Direction régionale",
    "Contrôle de gestion",
    "DSI",
    "Qualité"
]

EQUIPEMENTS = [
    "Four",
    "Chaudière",
    "Turbine",
    "Convoyeur",
    "Pont roulant",
    "Filtre",
    "Broyeur",
    "Système vapeur"
]


# =====================================================
# FONCTIONS UTILITAIRES ROBUSTES
# =====================================================

def random_date(start="2022-01-01", end="2026-12-31"):
    start_date = pd.to_datetime(start)
    end_date = pd.to_datetime(end)
    delta_days = (end_date - start_date).days
    return start_date + timedelta(days=random.randint(0, delta_days))


def save_excel(df, filename):
    path = os.path.join(OUTPUT_DIR, filename)
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"Fichier généré : {path} | {len(df)} lignes")


def introduce_missing_values(df, columns, rate=0.02):
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype("object")
            mask = np.random.rand(len(df)) < rate
            df.loc[mask, col] = np.nan
    return df


def introduce_duplicates(df, rate=0.01):
    n_duplicates = int(len(df) * rate)
    if n_duplicates <= 0:
        return df

    duplicates = df.sample(n=n_duplicates, random_state=SEED)
    df = pd.concat([df, duplicates], ignore_index=True)
    df = df.sample(frac=1, random_state=SEED).reset_index(drop=True)

    return df.head(N_ROWS)


def introduce_bad_dates(df, column, rate=0.03):
    """
    Ajoute volontairement des dates sales :
    - formats différents
    - textes
    - valeurs à confirmer

    Correction importante :
    la colonne est convertie en object pour autoriser le mélange date + texte.
    """
    if column not in df.columns:
        return df

    df[column] = df[column].astype("object")

    mask = np.random.rand(len(df)) < rate
    indexes = df.index[mask]

    def bad_format(value):
        if pd.isna(value):
            return value

        try:
            value = pd.to_datetime(value, errors="coerce")

            if pd.isna(value):
                return random.choice(["date inconnue", "à confirmer"])

            return random.choice([
                value.strftime("%d/%m/%Y"),
                value.strftime("%Y%m%d"),
                value.strftime("%m-%d-%Y"),
                "date inconnue",
                "à confirmer"
            ])

        except Exception:
            return random.choice(["date inconnue", "à confirmer"])

    for idx in indexes:
        df.at[idx, column] = bad_format(df.at[idx, column])

    return df


def introduce_text_amounts(df, column, rate=0.03):
    """
    Ajoute volontairement des montants sales :
    1250.50 devient par exemple '1 250,50 €'.

    Correction importante :
    la colonne est convertie en object pour autoriser le mélange nombre + texte.
    """
    if column not in df.columns:
        return df

    df[column] = df[column].astype("object")

    mask = np.random.rand(len(df)) < rate
    indexes = df.index[mask]

    def to_euro_text(value):
        if pd.isna(value):
            return value

        try:
            return f"{float(value):,.2f} €".replace(",", " ").replace(".", ",")
        except Exception:
            return value

    for idx in indexes:
        df.at[idx, column] = to_euro_text(df.at[idx, column])

    return df


def random_client_code():
    return f"CLI-{random.randint(1, N_ROWS):06d}"


def random_contract_code():
    return f"CTR-{random.randint(1, N_ROWS):07d}"


# =====================================================
# 1. CLIENTS
# =====================================================

def generate_clients():
    names = [
        "Métropole Rouen Normandie",
        "Caen la Mer",
        "Le Havre Seine Métropole",
        "Industrie Normande Services",
        "Papeteries de Normandie",
        "Agro Industries Manche",
        "Collectivité Côte Fleurie",
        "Syndicat Mixte Déchets Eure"
    ]

    rows = []

    for i in range(1, N_ROWS + 1):
        rows.append({
            "CodeClient": f"CLI-{i:06d}",
            "NomClient": f"{random.choice(names)} {random.randint(1, 999)}",
            "TypeClient": random.choice(TYPES_CLIENT),
            "Region": "Normandie",
            "Departement": random.choice(DEPARTEMENTS),
            "Ville": random.choice(VILLES),
            "SecteurActivite": random.choice([
                "Collectivité", "Industrie", "Santé", "Agroalimentaire",
                "Transport", "Administration", "BTP"
            ]),
            "StatutClient": random.choice(STATUTS_CLIENT),
            "DateCreation": random_date("2018-01-01", "2026-01-01"),
            "ResponsableCommercial": random.choice(RESPONSABLES),
            "EmailContact": f"contact{i}@client-normandie.fr",
            "Telephone": f"02{random.randint(10000000, 99999999)}",
            "SourceOrigine": random.choice(["ERP", "Excel", "CRM", "Import manuel"]),
            "CommentaireLibre": random.choice([
                "", "Client stratégique", "À vérifier", "Ancien fichier Excel",
                "Doublon potentiel", "Données incomplètes"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["CodeClient", "NomClient", "Departement", "EmailContact"],
        rate=0.015
    )

    df = introduce_duplicates(df, rate=0.015)
    df = introduce_bad_dates(df, "DateCreation", rate=0.025)

    mask = np.random.rand(len(df)) < 0.025
    df.loc[mask, "Region"] = np.random.choice(
        ["NORM", "Normandi", "normandie", "NORMANDIE ", "Normandie"],
        size=mask.sum()
    )

    mask = np.random.rand(len(df)) < 0.03
    df.loc[mask, "StatutClient"] = np.random.choice(
        ["actif", "ACTIF", "Résiliéé", "Suspendu ", "inactif"],
        size=mask.sum()
    )

    mask = np.random.rand(len(df)) < 0.01
    df.loc[mask, "EmailContact"] = np.random.choice(
        ["email_invalide", "contact@", "sans_email", "test.fr"],
        size=mask.sum()
    )

    save_excel(df, "01_Clients_Collectivites_Industriels.xlsx")
    return df


# =====================================================
# 2. CONTRATS
# =====================================================

def generate_contrats():
    rows = []

    for i in range(1, N_ROWS + 1):
        date_debut = random_date("2020-01-01", "2026-01-01")
        date_fin = date_debut + pd.DateOffset(months=random.choice([12, 24, 36, 48, 60]))
        type_contrat = random.choice(TYPES_CONTRAT)

        rows.append({
            "NumContrat": f"CTR-{i:07d}",
            "CodeClient": random_client_code(),
            "TypeContrat": type_contrat,
            "DateDebut": date_debut,
            "DateFin": date_fin,
            "StatutContrat": random.choice(STATUTS_CONTRAT),
            "ModeFacturation": random.choice(["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle"]),
            "PrixTonne": round(random.uniform(30, 250), 2),
            "PrixForfaitMensuel": round(random.uniform(500, 15000), 2),
            "TypePrestation": type_contrat,
            "ResponsableContrat": random.choice(RESPONSABLES),
            "SitePrincipal": random.choice(SITES),
            "AncienFichierExcel": random.choice([
                "Contrats_Clients_Normandie.xlsx",
                "Suivi_Contrats_V12_final.xlsx",
                "Contrats_SUEZ_reporting.xlsx"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["NumContrat", "CodeClient", "DateDebut", "PrixTonne"],
        rate=0.015
    )

    df = introduce_duplicates(df, rate=0.01)

    # Contrats actifs déjà expirés
    df["DateFin"] = df["DateFin"].astype("object")
    mask = np.random.rand(len(df)) < 0.03
    df.loc[mask, "StatutContrat"] = "Actif"
    df.loc[mask, "DateFin"] = pd.to_datetime("2023-12-31")

    # Prix négatifs
    mask = np.random.rand(len(df)) < 0.01
    df.loc[mask, "PrixTonne"] = -abs(pd.to_numeric(df.loc[mask, "PrixTonne"], errors="coerce"))

    # Montants texte
    df = introduce_text_amounts(df, "PrixForfaitMensuel", rate=0.03)

    # Statuts sales
    mask = np.random.rand(len(df)) < 0.035
    df.loc[mask, "StatutContrat"] = np.random.choice(
        ["actif", "ACTIF", "En cours", "resilie", "Résiliéé"],
        size=mask.sum()
    )

    # Dates sales à la fin pour éviter les erreurs de calcul
    df = introduce_bad_dates(df, "DateDebut", rate=0.02)
    df = introduce_bad_dates(df, "DateFin", rate=0.02)

    save_excel(df, "02_Contrats_Suez_Normandie.xlsx")
    return df


# =====================================================
# 3. TONNAGES DÉCHETS
# =====================================================

def generate_tonnages():
    rows = []

    for i in range(1, N_ROWS + 1):
        collecte = round(np.random.gamma(shape=4, scale=8), 2)
        traite = round(collecte * random.uniform(0.80, 1.05), 2)
        recycle = round(traite * random.uniform(0.10, 0.45), 2)
        valorise = round(traite * random.uniform(0.20, 0.70), 2)
        refus = round(max(traite - recycle - valorise, 0), 2)

        rows.append({
            "IdFlux": f"FLUX-{i:08d}",
            "DateFlux": random_date("2022-01-01", "2026-03-31"),
            "CodeClient": random_client_code(),
            "NumContrat": random_contract_code(),
            "SiteTraitement": random.choice(SITES),
            "TypeDechet": random.choice(TYPES_DECHET),
            "FamilleDechet": random.choice([
                "Déchet recyclable",
                "Déchet valorisable",
                "Déchet non recyclable",
                "Déchet dangereux",
                "Biodéchet"
            ]),
            "TonnageCollecte": collecte,
            "TonnageTraite": traite,
            "TonnageRecycle": recycle,
            "TonnageValorise": valorise,
            "TonnageRefus": refus,
            "ModeTransport": random.choice([
                "Camion benne",
                "Semi-remorque",
                "Benne ampliroll",
                "Transport interne",
                "Prestataire externe"
            ]),
            "OrigineFlux": random.choice([
                "Collecte municipale",
                "Industrie",
                "Déchetterie",
                "Site interne",
                "Prestataire"
            ]),
            "AncienFichierExcel": random.choice([
                "Suivi_Tonnages_Journalier.xlsx",
                "Tonnages_Exploitation_Final.xlsx",
                "Reporting_Dechets_Normandie.xlsx"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["IdFlux", "DateFlux", "CodeClient", "NumContrat", "SiteTraitement"],
        rate=0.015
    )

    df = introduce_duplicates(df, rate=0.01)

    for col in ["TonnageCollecte", "TonnageTraite", "TonnageValorise"]:
        mask = np.random.rand(len(df)) < 0.01
        df.loc[mask, col] = -abs(pd.to_numeric(df.loc[mask, col], errors="coerce"))

    mask = np.random.rand(len(df)) < 0.025
    df.loc[mask, "TonnageValorise"] = (
        pd.to_numeric(df.loc[mask, "TonnageTraite"], errors="coerce") * random.uniform(1.2, 2.5)
    )

    mask = np.random.rand(len(df)) < 0.02
    df.loc[mask, "TonnageTraite"] = (
        pd.to_numeric(df.loc[mask, "TonnageCollecte"], errors="coerce") * random.uniform(1.3, 3.0)
    )

    mask = np.random.rand(len(df)) < 0.02
    df.loc[mask, "TypeDechet"] = np.random.choice(
        ["OM", "ordures menageres", "DIB", "plastique", "METAL", "dechet inconnu"],
        size=mask.sum()
    )

    df = introduce_bad_dates(df, "DateFlux", rate=0.03)

    save_excel(df, "03_Tonnages_Dechets.xlsx")
    return df


# =====================================================
# 4. VALORISATION ÉNERGÉTIQUE
# =====================================================

def generate_valorisation_energetique():
    rows = []

    for i in range(1, N_ROWS + 1):
        tonnage = round(np.random.gamma(shape=5, scale=10), 2)
        mwh_elec = round(tonnage * random.uniform(0.35, 0.85), 2)
        mwh_chaleur = round(tonnage * random.uniform(0.50, 1.20), 2)
        mwh_vendus = round((mwh_elec + mwh_chaleur) * random.uniform(0.60, 0.95), 2)
        heures_fonctionnement = round(random.uniform(10, 24), 2)
        heures_arret = round(max(24 - heures_fonctionnement, 0), 2)

        rows.append({
            "IdProduction": f"PROD-{i:08d}",
            "DateProduction": random_date("2022-01-01", "2026-03-31"),
            "SiteUVE": random.choice([
                "UVE Caen Colombelles",
                "Centre de valorisation Évreux",
                "Plateforme déchets Le Havre"
            ]),
            "TonnageValorise": tonnage,
            "MWhElectriciteProduite": mwh_elec,
            "MWhChaleurProduite": mwh_chaleur,
            "MWhVendus": mwh_vendus,
            "RendementEnergetique": round(random.uniform(0.45, 0.90), 4),
            "HeuresFonctionnement": heures_fonctionnement,
            "HeuresArret": heures_arret,
            "CauseArret": random.choice([
                "Aucun",
                "Maintenance planifiée",
                "Panne équipement",
                "Surchauffe",
                "Incident sécurité",
                "Arrêt réglementaire",
                "Défaut capteur"
            ]),
            "StatutProduction": random.choice(STATUTS_PRODUCTION),
            "AncienFichierExcel": random.choice([
                "Production_Energie_UVE.xlsx",
                "Suivi_Energie_Normandie.xlsx",
                "Valorisation_Energetique_Final.xlsx"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(df, ["IdProduction", "DateProduction", "SiteUVE"], rate=0.015)
    df = introduce_duplicates(df, rate=0.01)

    mask = np.random.rand(len(df)) < 0.02
    df.loc[mask, "TonnageValorise"] = 0
    df.loc[mask, "MWhElectriciteProduite"] = np.random.uniform(10, 100, size=mask.sum())

    mask = np.random.rand(len(df)) < 0.015
    df.loc[mask, "RendementEnergetique"] = np.random.uniform(1.1, 2.5, size=mask.sum())

    mask = np.random.rand(len(df)) < 0.015
    df.loc[mask, "HeuresFonctionnement"] = np.random.uniform(25, 40, size=mask.sum())

    mask = np.random.rand(len(df)) < 0.03
    df.loc[mask, "StatutProduction"] = np.random.choice(
        ["normal", "NORMAL", "arret", "incident ", "maintenance"],
        size=mask.sum()
    )

    df = introduce_bad_dates(df, "DateProduction", rate=0.03)

    save_excel(df, "04_Valorisation_Energetique.xlsx")
    return df


# =====================================================
# 5. FACTURATION
# =====================================================

def generate_facturation():
    rows = []

    for i in range(1, N_ROWS + 1):
        date_facture = random_date("2022-01-01", "2026-03-31")
        date_echeance = date_facture + pd.DateOffset(days=random.choice([15, 30, 45, 60]))

        montant_ht = round(random.uniform(500, 80000), 2)
        tva = round(montant_ht * 0.20, 2)
        montant_ttc = round(montant_ht + tva, 2)

        statut = random.choice(STATUTS_FACTURE)

        if statut == "Payée":
            date_paiement = date_facture + pd.DateOffset(days=random.randint(1, 60))
        else:
            date_paiement = None

        rows.append({
            "NumFacture": f"FAC-{i:08d}",
            "CodeClient": random_client_code(),
            "NumContrat": random_contract_code(),
            "DateFacture": date_facture,
            "DateEcheance": date_echeance,
            "MontantHT": montant_ht,
            "TVA": tva,
            "MontantTTC": montant_ttc,
            "StatutPaiement": statut,
            "DatePaiement": date_paiement,
            "TypePrestation": random.choice(TYPES_CONTRAT),
            "CompteComptable": random.choice(["706000", "706100", "706200", "411000", "704000"]),
            "AncienFichierExcel": random.choice([
                "Suivi_Facturation_Normandie.xlsx",
                "Impayes_Clients_Final.xlsx",
                "Factures_SUEZ_Reporting.xlsx"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["NumFacture", "CodeClient", "NumContrat", "DateFacture"],
        rate=0.015
    )

    df = introduce_duplicates(df, rate=0.01)

    mask = np.random.rand(len(df)) < 0.01
    df.loc[mask, "MontantTTC"] = -abs(pd.to_numeric(df.loc[mask, "MontantTTC"], errors="coerce"))

    mask = np.random.rand(len(df)) < 0.025
    df.loc[mask, "StatutPaiement"] = "Payée"
    df.loc[mask, "DatePaiement"] = None

    mask = np.random.rand(len(df)) < 0.02
    df.loc[mask, "StatutPaiement"] = "En retard"
    df.loc[mask, "DatePaiement"] = pd.to_datetime("2026-01-15")

    # Échéance avant facture AVANT de salir les dates
    df["DateEcheance"] = df["DateEcheance"].astype("object")
    mask = np.random.rand(len(df)) < 0.015
    valid_dates = pd.to_datetime(df.loc[mask, "DateFacture"], errors="coerce")
    df.loc[mask, "DateEcheance"] = valid_dates - pd.DateOffset(days=10)

    mask = np.random.rand(len(df)) < 0.035
    df.loc[mask, "StatutPaiement"] = np.random.choice(
        ["payee", "PAYÉE", "Retard", "en attente ", "litige client"],
        size=mask.sum()
    )

    df = introduce_text_amounts(df, "MontantHT", rate=0.025)
    df = introduce_text_amounts(df, "MontantTTC", rate=0.035)

    df = introduce_bad_dates(df, "DateFacture", rate=0.025)
    df = introduce_bad_dates(df, "DateEcheance", rate=0.02)
    df = introduce_bad_dates(df, "DatePaiement", rate=0.02)

    save_excel(df, "05_Facturation_Prestations.xlsx")
    return df


# =====================================================
# 6. MAINTENANCE
# =====================================================

def generate_interventions():
    rows = []

    for i in range(1, N_ROWS + 1):
        date_intervention = random_date("2022-01-01", "2026-03-31")
        statut = random.choice(STATUTS_INTERVENTION)

        if statut == "Clôturée":
            date_cloture = date_intervention + pd.DateOffset(days=random.randint(0, 15))
        else:
            date_cloture = None

        duree = round(random.uniform(0.5, 12), 2)
        cout = round(duree * random.uniform(80, 250), 2)

        rows.append({
            "NumIntervention": f"INT-{i:08d}",
            "SiteTraitement": random.choice(SITES),
            "Equipement": random.choice(EQUIPEMENTS),
            "DateIntervention": date_intervention,
            "DateCloture": date_cloture,
            "Technicien": f"TECH-{random.randint(1, 50):03d}",
            "TypeIntervention": random.choice([
                "Maintenance préventive",
                "Maintenance corrective",
                "Dépannage",
                "Inspection sécurité",
                "Remplacement équipement",
                "Arrêt technique"
            ]),
            "StatutIntervention": statut,
            "DureeHeures": duree,
            "CoutIntervention": cout,
            "Criticite": random.choice(["Faible", "Moyenne", "Haute", "Critique"]),
            "Commentaire": random.choice([
                "", "Pièce manquante", "Arrêt ligne nécessaire",
                "Intervention urgente", "Contrôle à replanifier",
                "Anomalie capteur", "Problème résolu"
            ]),
            "AncienFichierExcel": random.choice([
                "Maintenance_UVE.xlsx",
                "Suivi_Interventions_Final.xlsx",
                "Arrets_Equipements_Normandie.xlsx"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["NumIntervention", "SiteTraitement", "Equipement", "DateIntervention", "Technicien"],
        rate=0.015
    )

    df = introduce_duplicates(df, rate=0.01)

    mask = np.random.rand(len(df)) < 0.025
    df.loc[mask, "StatutIntervention"] = "Clôturée"
    df.loc[mask, "DateCloture"] = None

    mask = np.random.rand(len(df)) < 0.01
    df.loc[mask, "DureeHeures"] = -abs(pd.to_numeric(df.loc[mask, "DureeHeures"], errors="coerce"))

    mask = np.random.rand(len(df)) < 0.035
    df.loc[mask, "StatutIntervention"] = np.random.choice(
        ["cloturee", "CLOTUREE", "annulee", "En retard ", "planifiee"],
        size=mask.sum()
    )

    df = introduce_text_amounts(df, "CoutIntervention", rate=0.03)

    df = introduce_bad_dates(df, "DateIntervention", rate=0.025)
    df = introduce_bad_dates(df, "DateCloture", rate=0.02)

    save_excel(df, "06_Interventions_Maintenance.xlsx")
    return df


# =====================================================
# 7. SUIVI MIGRATION EXCEL
# =====================================================

def generate_suivi_migration_excel():
    fichiers_base = [
        "Suivi_Tonnages_Journalier.xlsx",
        "Production_Energie_UVE.xlsx",
        "Suivi_Facturation_Normandie.xlsx",
        "Maintenance_UVE.xlsx",
        "Contrats_Clients_Normandie.xlsx",
        "Suivi_Impayes_Clients.xlsx",
        "Reporting_Direction_Regionale.xlsx",
        "Controle_Qualite_Dechets.xlsx",
        "Planning_Interventions.xlsx",
        "Suivi_Arrets_Equipements.xlsx"
    ]

    rows = []

    for i in range(1, N_ROWS + 1):
        rows.append({
            "IdSuivi": f"MIG-{i:08d}",
            "NomFichierExcel": random.choice(fichiers_base),
            "Service": random.choice(SERVICES),
            "Proprietaire": random.choice(RESPONSABLES),
            "FrequenceUsage": random.choice(["Quotidienne", "Hebdomadaire", "Mensuelle", "Trimestrielle"]),
            "Criticite": random.choice(["Faible", "Moyenne", "Haute", "Critique"]),
            "RisqueMetier": random.choice([
                "Erreur reporting",
                "Mauvaise facturation",
                "Mauvais suivi tonnage",
                "Erreur KPI direction",
                "Perte de traçabilité",
                "Double saisie",
                "Décision basée sur données obsolètes"
            ]),
            "Cible": random.choice([
                "Business Central",
                "Power BI",
                "Microsoft Fabric",
                "Business Central + Power BI",
                "Fabric + Power BI",
                "À supprimer"
            ]),
            "StatutMigration": random.choice([
                "Non démarré",
                "En analyse",
                "En nettoyage",
                "En mapping",
                "En recette",
                "Migré",
                "Rejeté"
            ]),
            "DateCible": random_date("2026-04-01", "2026-12-31"),
            "NombreUtilisateurs": random.randint(1, 80),
            "TempsRetraitementHeures": round(random.uniform(0.5, 16), 2),
            "Commentaire": random.choice([
                "", "Fichier critique", "Macros à analyser",
                "TCD complexes", "Doublons fréquents",
                "À remplacer par Power BI", "Validation métier nécessaire"
            ])
        })

    df = pd.DataFrame(rows)

    df = introduce_missing_values(
        df,
        ["NomFichierExcel", "Service", "Proprietaire", "Cible"],
        rate=0.02
    )

    df = introduce_duplicates(df, rate=0.03)

    mask = np.random.rand(len(df)) < 0.04
    df.loc[mask, "StatutMigration"] = np.random.choice(
        ["migré", "MIGRE", "en cours", "A faire", "abandonné"],
        size=mask.sum()
    )

    mask = np.random.rand(len(df)) < 0.035
    df.loc[mask, "Criticite"] = np.random.choice(
        ["critique", "CRITIQUE", "Haute ", "moyen", "urgent"],
        size=mask.sum()
    )

    mask = np.random.rand(len(df)) < 0.01
    df.loc[mask, "NombreUtilisateurs"] = -abs(
        pd.to_numeric(df.loc[mask, "NombreUtilisateurs"], errors="coerce")
    )

    df = introduce_bad_dates(df, "DateCible", rate=0.025)

    save_excel(df, "07_Suivi_Migration_Excel.xlsx")
    return df


# =====================================================
# EXÉCUTION
# =====================================================

def main():
    print("Génération des fichiers Excel sales du projet énergie / environnement...")

    generate_clients()
    generate_contrats()
    generate_tonnages()
    generate_valorisation_energetique()
    generate_facturation()
    generate_interventions()
    generate_suivi_migration_excel()

    print("\nGénération terminée.")
    print(f"Dossier de sortie : {OUTPUT_DIR}")
    print("\nFichiers générés :")

    for file in os.listdir(OUTPUT_DIR):
        print(f"- {file}")


if __name__ == "__main__":
    main()