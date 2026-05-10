import os
import re
from datetime import datetime

import numpy as np
import pandas as pd


# =====================================================
# PARAMÈTRES
# =====================================================

INPUT_DIR = "suez_normandie_dirty_files"
OUTPUT_DIR = "suez_normandie_audit_outputs"

os.makedirs(OUTPUT_DIR, exist_ok=True)

FILES = {
    "clients": "01_Clients_Collectivites_Industriels.xlsx",
    "contrats": "02_Contrats_Suez_Normandie.xlsx",
    "tonnages": "03_Tonnages_Dechets.xlsx",
    "energie": "04_Valorisation_Energetique.xlsx",
    "facturation": "05_Facturation_Prestations.xlsx",
    "maintenance": "06_Interventions_Maintenance.xlsx",
    "migration_excel": "07_Suivi_Migration_Excel.xlsx",
}


# =====================================================
# FONCTIONS UTILITAIRES
# =====================================================

def load_excel_file(filename):
    path = os.path.join(INPUT_DIR, filename)

    if not os.path.exists(path):
        raise FileNotFoundError(f"Fichier introuvable : {path}")

    return pd.read_excel(path, dtype="object")


def normalize_text(value):
    if pd.isna(value):
        return np.nan

    return str(value).strip()


def parse_amount(value):
    """
    Convertit les montants sales en nombre.
    Exemples acceptés :
    - 1200.50
    - "1 200,50 €"
    - "1200,50"
    """
    if pd.isna(value):
        return np.nan

    value = str(value).strip()

    value = value.replace("€", "")
    value = value.replace(" ", "")
    value = value.replace(",", ".")

    try:
        return float(value)
    except Exception:
        return np.nan


def parse_date(value):
    """
    Convertit les dates sales en datetime.
    Les valeurs comme 'date inconnue' ou 'à confirmer' deviennent NaT.
    """
    if pd.isna(value):
        return pd.NaT

    try:
        return pd.to_datetime(value, errors="coerce", dayfirst=True)
    except Exception:
        return pd.NaT


def count_missing(df):
    rows = []

    for col in df.columns:
        missing_count = df[col].isna().sum()
        missing_rate = missing_count / len(df) if len(df) > 0 else 0

        rows.append({
            "Colonne": col,
            "Nb_Lignes": len(df),
            "Nb_Valeurs_Manquantes": missing_count,
            "Taux_Valeurs_Manquantes": round(missing_rate, 4),
        })

    return pd.DataFrame(rows)


def count_duplicates(df, subset_cols=None):
    if subset_cols:
        existing_cols = [col for col in subset_cols if col in df.columns]

        if not existing_cols:
            return 0

        return df.duplicated(subset=existing_cols).sum()

    return df.duplicated().sum()


def audit_basic_structure(domain, df):
    return {
        "Domaine": domain,
        "Nb_Lignes": len(df),
        "Nb_Colonnes": len(df.columns),
        "Nb_Doublons_Complets": int(df.duplicated().sum()),
        "Nb_Cellules_Manquantes": int(df.isna().sum().sum()),
        "Date_Audit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def build_anomaly(domain, rule, severity, count, comment):
    return {
        "Domaine": domain,
        "Regle_Controle": rule,
        "Criticite": severity,
        "Nb_Anomalies": int(count),
        "Commentaire": comment,
    }


# =====================================================
# AUDIT CLIENTS
# =====================================================

def audit_clients(df):
    anomalies = []

    required_cols = ["CodeClient", "NomClient", "EmailContact", "Region", "StatutClient"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Clients",
                    f"Valeurs manquantes sur {col}",
                    "Critique" if col in ["CodeClient", "NomClient"] else "Majeure",
                    df[col].isna().sum(),
                    f"La colonne {col} ne doit pas contenir de valeurs manquantes."
                )
            )

    if "CodeClient" in df.columns:
        anomalies.append(
            build_anomaly(
                "Clients",
                "Doublons sur CodeClient",
                "Critique",
                df.duplicated(subset=["CodeClient"]).sum(),
                "Un client doit avoir un identifiant unique."
            )
        )

    if "Region" in df.columns:
        region_clean = df["Region"].astype(str).str.strip().str.lower()
        count = (~region_clean.isin(["normandie"])).sum()
        anomalies.append(
            build_anomaly(
                "Clients",
                "Région différente de Normandie",
                "Majeure",
                count,
                "La région doit être normalisée à 'Normandie'."
            )
        )

    if "EmailContact" in df.columns:
        email_pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        emails = df["EmailContact"].fillna("").astype(str).str.strip()
        invalid_email_count = (~emails.str.match(email_pattern)).sum()

        anomalies.append(
            build_anomaly(
                "Clients",
                "Emails invalides",
                "Mineure",
                invalid_email_count,
                "Les emails doivent respecter un format valide."
            )
        )

    if "StatutClient" in df.columns:
        valid_status = ["Actif", "Inactif", "Suspendu", "Résilié"]
        status = df["StatutClient"].fillna("").astype(str).str.strip()
        count = (~status.isin(valid_status)).sum()

        anomalies.append(
            build_anomaly(
                "Clients",
                "Statuts clients non normalisés",
                "Majeure",
                count,
                "Les statuts clients doivent appartenir à la liste officielle."
            )
        )

    return anomalies


# =====================================================
# AUDIT CONTRATS
# =====================================================

def audit_contrats(df):
    anomalies = []

    required_cols = ["NumContrat", "CodeClient", "DateDebut", "DateFin", "StatutContrat"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Contrats",
                    f"Valeurs manquantes sur {col}",
                    "Critique" if col in ["NumContrat", "CodeClient", "DateDebut"] else "Majeure",
                    df[col].isna().sum(),
                    f"La colonne {col} est nécessaire pour le mapping Business Central."
                )
            )

    if "NumContrat" in df.columns:
        anomalies.append(
            build_anomaly(
                "Contrats",
                "Doublons sur NumContrat",
                "Critique",
                df.duplicated(subset=["NumContrat"]).sum(),
                "Un contrat doit avoir un numéro unique."
            )
        )

    if "DateDebut" in df.columns:
        parsed = df["DateDebut"].apply(parse_date)
        anomalies.append(
            build_anomaly(
                "Contrats",
                "Dates de début invalides",
                "Critique",
                parsed.isna().sum(),
                "Les dates de début doivent être convertibles en date."
            )
        )

    if "DateFin" in df.columns:
        parsed = df["DateFin"].apply(parse_date)
        anomalies.append(
            build_anomaly(
                "Contrats",
                "Dates de fin invalides",
                "Majeure",
                parsed.isna().sum(),
                "Les dates de fin doivent être convertibles en date."
            )
        )

    if {"DateDebut", "DateFin"}.issubset(df.columns):
        date_debut = df["DateDebut"].apply(parse_date)
        date_fin = df["DateFin"].apply(parse_date)

        count = ((date_fin < date_debut) & date_debut.notna() & date_fin.notna()).sum()

        anomalies.append(
            build_anomaly(
                "Contrats",
                "Date de fin antérieure à la date de début",
                "Critique",
                count,
                "La date de fin ne doit pas être antérieure à la date de début."
            )
        )

    if {"StatutContrat", "DateFin"}.issubset(df.columns):
        status = df["StatutContrat"].fillna("").astype(str).str.strip()
        date_fin = df["DateFin"].apply(parse_date)

        count = ((status == "Actif") & date_fin.notna() & (date_fin < pd.Timestamp.today())).sum()

        anomalies.append(
            build_anomaly(
                "Contrats",
                "Contrats actifs avec date de fin passée",
                "Majeure",
                count,
                "Un contrat actif ne devrait pas avoir une date de fin déjà dépassée."
            )
        )

    for amount_col in ["PrixTonne", "PrixForfaitMensuel"]:
        if amount_col in df.columns:
            amount = df[amount_col].apply(parse_amount)

            anomalies.append(
                build_anomaly(
                    "Contrats",
                    f"Montants invalides sur {amount_col}",
                    "Majeure",
                    amount.isna().sum(),
                    f"La colonne {amount_col} doit être convertible en nombre."
                )
            )

            anomalies.append(
                build_anomaly(
                    "Contrats",
                    f"Montants négatifs sur {amount_col}",
                    "Critique",
                    (amount < 0).sum(),
                    f"La colonne {amount_col} ne doit pas contenir de valeurs négatives."
                )
            )

    return anomalies


# =====================================================
# AUDIT TONNAGES
# =====================================================

def audit_tonnages(df):
    anomalies = []

    required_cols = ["IdFlux", "DateFlux", "CodeClient", "NumContrat", "SiteTraitement"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Tonnages",
                    f"Valeurs manquantes sur {col}",
                    "Critique",
                    df[col].isna().sum(),
                    f"La colonne {col} est nécessaire au suivi des flux déchets."
                )
            )

    if "IdFlux" in df.columns:
        anomalies.append(
            build_anomaly(
                "Tonnages",
                "Doublons sur IdFlux",
                "Critique",
                df.duplicated(subset=["IdFlux"]).sum(),
                "Chaque flux doit avoir un identifiant unique."
            )
        )

    if "DateFlux" in df.columns:
        parsed = df["DateFlux"].apply(parse_date)
        anomalies.append(
            build_anomaly(
                "Tonnages",
                "Dates de flux invalides",
                "Critique",
                parsed.isna().sum(),
                "Les dates de flux doivent être convertibles en date."
            )
        )

    tonnage_cols = [
        "TonnageCollecte",
        "TonnageTraite",
        "TonnageRecycle",
        "TonnageValorise",
        "TonnageRefus"
    ]

    parsed_tonnages = {}

    for col in tonnage_cols:
        if col in df.columns:
            parsed_tonnages[col] = df[col].apply(parse_amount)

            anomalies.append(
                build_anomaly(
                    "Tonnages",
                    f"Tonnages invalides sur {col}",
                    "Majeure",
                    parsed_tonnages[col].isna().sum(),
                    f"La colonne {col} doit être numérique."
                )
            )

            anomalies.append(
                build_anomaly(
                    "Tonnages",
                    f"Tonnages négatifs sur {col}",
                    "Critique",
                    (parsed_tonnages[col] < 0).sum(),
                    f"La colonne {col} ne doit pas contenir de valeurs négatives."
                )
            )

    if {"TonnageValorise", "TonnageTraite"}.issubset(parsed_tonnages):
        count = (
            parsed_tonnages["TonnageValorise"]
            > parsed_tonnages["TonnageTraite"]
        ).sum()

        anomalies.append(
            build_anomaly(
                "Tonnages",
                "Tonnage valorisé supérieur au tonnage traité",
                "Majeure",
                count,
                "Le tonnage valorisé ne doit normalement pas dépasser le tonnage traité."
            )
        )

    if {"TonnageTraite", "TonnageCollecte"}.issubset(parsed_tonnages):
        count = (
            parsed_tonnages["TonnageTraite"]
            > parsed_tonnages["TonnageCollecte"] * 1.2
        ).sum()

        anomalies.append(
            build_anomaly(
                "Tonnages",
                "Tonnage traité très supérieur au tonnage collecté",
                "Majeure",
                count,
                "Écart à valider avec les équipes exploitation."
            )
        )

    return anomalies


# =====================================================
# AUDIT ÉNERGIE
# =====================================================

def audit_energie(df):
    anomalies = []

    required_cols = ["IdProduction", "DateProduction", "SiteUVE"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Énergie",
                    f"Valeurs manquantes sur {col}",
                    "Critique",
                    df[col].isna().sum(),
                    f"La colonne {col} est obligatoire pour le suivi de production énergétique."
                )
            )

    if "IdProduction" in df.columns:
        anomalies.append(
            build_anomaly(
                "Énergie",
                "Doublons sur IdProduction",
                "Critique",
                df.duplicated(subset=["IdProduction"]).sum(),
                "Chaque production doit avoir un identifiant unique."
            )
        )

    if "DateProduction" in df.columns:
        parsed = df["DateProduction"].apply(parse_date)
        anomalies.append(
            build_anomaly(
                "Énergie",
                "Dates de production invalides",
                "Critique",
                parsed.isna().sum(),
                "Les dates de production doivent être convertibles en date."
            )
        )

    numeric_cols = [
        "TonnageValorise",
        "MWhElectriciteProduite",
        "MWhChaleurProduite",
        "MWhVendus",
        "RendementEnergetique",
        "HeuresFonctionnement",
        "HeuresArret"
    ]

    parsed = {}

    for col in numeric_cols:
        if col in df.columns:
            parsed[col] = df[col].apply(parse_amount)

            anomalies.append(
                build_anomaly(
                    "Énergie",
                    f"Valeurs numériques invalides sur {col}",
                    "Majeure",
                    parsed[col].isna().sum(),
                    f"La colonne {col} doit être numérique."
                )
            )

    if {"TonnageValorise", "MWhElectriciteProduite"}.issubset(parsed):
        count = (
            (parsed["TonnageValorise"] == 0)
            & (parsed["MWhElectriciteProduite"] > 0)
        ).sum()

        anomalies.append(
            build_anomaly(
                "Énergie",
                "MWh produits avec tonnage valorisé égal à 0",
                "Critique",
                count,
                "Une production énergétique sans tonnage valorisé doit être vérifiée."
            )
        )

    if "RendementEnergetique" in parsed:
        count = (parsed["RendementEnergetique"] > 1).sum()

        anomalies.append(
            build_anomaly(
                "Énergie",
                "Rendement énergétique supérieur à 100%",
                "Critique",
                count,
                "Un rendement supérieur à 100% est incohérent."
            )
        )

    if "HeuresFonctionnement" in parsed:
        count = (parsed["HeuresFonctionnement"] > 24).sum()

        anomalies.append(
            build_anomaly(
                "Énergie",
                "Heures de fonctionnement supérieures à 24h",
                "Critique",
                count,
                "Sur une journée, les heures de fonctionnement ne doivent pas dépasser 24."
            )
        )

    return anomalies


# =====================================================
# AUDIT FACTURATION
# =====================================================

def audit_facturation(df):
    anomalies = []

    required_cols = ["NumFacture", "CodeClient", "NumContrat", "DateFacture", "MontantTTC"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Facturation",
                    f"Valeurs manquantes sur {col}",
                    "Critique",
                    df[col].isna().sum(),
                    f"La colonne {col} est nécessaire pour le reporting financier."
                )
            )

    if "NumFacture" in df.columns:
        anomalies.append(
            build_anomaly(
                "Facturation",
                "Doublons sur NumFacture",
                "Critique",
                df.duplicated(subset=["NumFacture"]).sum(),
                "Chaque facture doit avoir un numéro unique."
            )
        )

    for date_col in ["DateFacture", "DateEcheance", "DatePaiement"]:
        if date_col in df.columns:
            parsed = df[date_col].apply(parse_date)

            anomalies.append(
                build_anomaly(
                    "Facturation",
                    f"Dates invalides sur {date_col}",
                    "Majeure",
                    parsed.isna().sum(),
                    f"La colonne {date_col} doit être convertible en date."
                )
            )

    for amount_col in ["MontantHT", "TVA", "MontantTTC"]:
        if amount_col in df.columns:
            amount = df[amount_col].apply(parse_amount)

            anomalies.append(
                build_anomaly(
                    "Facturation",
                    f"Montants invalides sur {amount_col}",
                    "Majeure",
                    amount.isna().sum(),
                    f"La colonne {amount_col} doit être numérique."
                )
            )

            anomalies.append(
                build_anomaly(
                    "Facturation",
                    f"Montants négatifs sur {amount_col}",
                    "Critique",
                    (amount < 0).sum(),
                    f"La colonne {amount_col} ne doit pas contenir de valeurs négatives."
                )
            )

    if {"StatutPaiement", "DatePaiement"}.issubset(df.columns):
        status = df["StatutPaiement"].fillna("").astype(str).str.strip()
        date_paiement = df["DatePaiement"].apply(parse_date)

        count = ((status == "Payée") & date_paiement.isna()).sum()

        anomalies.append(
            build_anomaly(
                "Facturation",
                "Factures payées sans date de paiement",
                "Critique",
                count,
                "Une facture payée doit avoir une date de paiement."
            )
        )

    if {"DateFacture", "DateEcheance"}.issubset(df.columns):
        date_facture = df["DateFacture"].apply(parse_date)
        date_echeance = df["DateEcheance"].apply(parse_date)

        count = (
            (date_echeance < date_facture)
            & date_facture.notna()
            & date_echeance.notna()
        ).sum()

        anomalies.append(
            build_anomaly(
                "Facturation",
                "Date d’échéance antérieure à la date facture",
                "Critique",
                count,
                "La date d’échéance ne doit pas être avant la date de facture."
            )
        )

    if "StatutPaiement" in df.columns:
        valid_status = ["Payée", "En attente", "En retard", "Partiellement payée", "Litige"]
        status = df["StatutPaiement"].fillna("").astype(str).str.strip()

        count = (~status.isin(valid_status)).sum()

        anomalies.append(
            build_anomaly(
                "Facturation",
                "Statuts de paiement non normalisés",
                "Majeure",
                count,
                "Les statuts de paiement doivent être normalisés."
            )
        )

    return anomalies


# =====================================================
# AUDIT MAINTENANCE
# =====================================================

def audit_maintenance(df):
    anomalies = []

    required_cols = ["NumIntervention", "SiteTraitement", "Equipement", "DateIntervention"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Maintenance",
                    f"Valeurs manquantes sur {col}",
                    "Critique",
                    df[col].isna().sum(),
                    f"La colonne {col} est nécessaire au suivi maintenance."
                )
            )

    if "NumIntervention" in df.columns:
        anomalies.append(
            build_anomaly(
                "Maintenance",
                "Doublons sur NumIntervention",
                "Critique",
                df.duplicated(subset=["NumIntervention"]).sum(),
                "Chaque intervention doit avoir un numéro unique."
            )
        )

    for date_col in ["DateIntervention", "DateCloture"]:
        if date_col in df.columns:
            parsed = df[date_col].apply(parse_date)

            anomalies.append(
                build_anomaly(
                    "Maintenance",
                    f"Dates invalides sur {date_col}",
                    "Majeure",
                    parsed.isna().sum(),
                    f"La colonne {date_col} doit être convertible en date."
                )
            )

    for amount_col in ["DureeHeures", "CoutIntervention"]:
        if amount_col in df.columns:
            amount = df[amount_col].apply(parse_amount)

            anomalies.append(
                build_anomaly(
                    "Maintenance",
                    f"Valeurs invalides sur {amount_col}",
                    "Majeure",
                    amount.isna().sum(),
                    f"La colonne {amount_col} doit être numérique."
                )
            )

            anomalies.append(
                build_anomaly(
                    "Maintenance",
                    f"Valeurs négatives sur {amount_col}",
                    "Critique",
                    (amount < 0).sum(),
                    f"La colonne {amount_col} ne doit pas contenir de valeurs négatives."
                )
            )

    if {"StatutIntervention", "DateCloture"}.issubset(df.columns):
        status = df["StatutIntervention"].fillna("").astype(str).str.strip()
        date_cloture = df["DateCloture"].apply(parse_date)

        count = ((status == "Clôturée") & date_cloture.isna()).sum()

        anomalies.append(
            build_anomaly(
                "Maintenance",
                "Interventions clôturées sans date de clôture",
                "Critique",
                count,
                "Une intervention clôturée doit avoir une date de clôture."
            )
        )

    if "StatutIntervention" in df.columns:
        valid_status = ["Planifiée", "En cours", "Clôturée", "Annulée", "En retard"]
        status = df["StatutIntervention"].fillna("").astype(str).str.strip()

        count = (~status.isin(valid_status)).sum()

        anomalies.append(
            build_anomaly(
                "Maintenance",
                "Statuts intervention non normalisés",
                "Majeure",
                count,
                "Les statuts doivent appartenir à la liste officielle."
            )
        )

    return anomalies


# =====================================================
# AUDIT MIGRATION EXCEL
# =====================================================

def audit_migration_excel(df):
    anomalies = []

    required_cols = ["IdSuivi", "NomFichierExcel", "Service", "Proprietaire", "Cible"]

    for col in required_cols:
        if col in df.columns:
            anomalies.append(
                build_anomaly(
                    "Migration Excel",
                    f"Valeurs manquantes sur {col}",
                    "Majeure",
                    df[col].isna().sum(),
                    f"La colonne {col} est nécessaire au suivi de migration Excel."
                )
            )

    if "IdSuivi" in df.columns:
        anomalies.append(
            build_anomaly(
                "Migration Excel",
                "Doublons sur IdSuivi",
                "Critique",
                df.duplicated(subset=["IdSuivi"]).sum(),
                "Chaque ligne de suivi doit avoir un identifiant unique."
            )
        )

    if "DateCible" in df.columns:
        parsed = df["DateCible"].apply(parse_date)

        anomalies.append(
            build_anomaly(
                "Migration Excel",
                "Dates cibles invalides",
                "Majeure",
                parsed.isna().sum(),
                "Les dates cibles doivent être convertibles en date."
            )
        )

    if "NombreUtilisateurs" in df.columns:
        users = df["NombreUtilisateurs"].apply(parse_amount)

        anomalies.append(
            build_anomaly(
                "Migration Excel",
                "Nombre utilisateurs négatif",
                "Critique",
                (users < 0).sum(),
                "Le nombre d’utilisateurs ne peut pas être négatif."
            )
        )

    if "Criticite" in df.columns:
        valid_values = ["Faible", "Moyenne", "Haute", "Critique"]
        criticite = df["Criticite"].fillna("").astype(str).str.strip()

        count = (~criticite.isin(valid_values)).sum()

        anomalies.append(
            build_anomaly(
                "Migration Excel",
                "Criticités non normalisées",
                "Majeure",
                count,
                "La criticité doit appartenir à la liste officielle."
            )
        )

    if "StatutMigration" in df.columns:
        valid_values = [
            "Non démarré",
            "En analyse",
            "En nettoyage",
            "En mapping",
            "En recette",
            "Migré",
            "Rejeté"
        ]

        status = df["StatutMigration"].fillna("").astype(str).str.strip()

        count = (~status.isin(valid_values)).sum()

        anomalies.append(
            build_anomaly(
                "Migration Excel",
                "Statuts de migration non normalisés",
                "Majeure",
                count,
                "Les statuts de migration doivent être harmonisés."
            )
        )

    return anomalies


# =====================================================
# AUDIT GLOBAL
# =====================================================

def main():
    print("Démarrage de l’audit qualité des fichiers Excel sales...")

    structure_rows = []
    missing_rows = []
    anomaly_rows = []

    loaded_data = {}

    for domain, filename in FILES.items():
        print(f"Lecture du fichier : {filename}")

        df = load_excel_file(filename)
        loaded_data[domain] = df

        structure_rows.append(audit_basic_structure(domain, df))

        missing_df = count_missing(df)
        missing_df.insert(0, "Domaine", domain)
        missing_rows.append(missing_df)

        if domain == "clients":
            anomaly_rows.extend(audit_clients(df))

        elif domain == "contrats":
            anomaly_rows.extend(audit_contrats(df))

        elif domain == "tonnages":
            anomaly_rows.extend(audit_tonnages(df))

        elif domain == "energie":
            anomaly_rows.extend(audit_energie(df))

        elif domain == "facturation":
            anomaly_rows.extend(audit_facturation(df))

        elif domain == "maintenance":
            anomaly_rows.extend(audit_maintenance(df))

        elif domain == "migration_excel":
            anomaly_rows.extend(audit_migration_excel(df))

    structure_df = pd.DataFrame(structure_rows)
    missing_report_df = pd.concat(missing_rows, ignore_index=True)
    anomalies_df = pd.DataFrame(anomaly_rows)

    anomalies_df = anomalies_df.sort_values(
        by=["Criticite", "Nb_Anomalies"],
        ascending=[True, False]
    )

    # Synthèse par domaine
    summary_by_domain = (
        anomalies_df
        .groupby("Domaine", as_index=False)
        .agg(
            Nb_Controles=("Regle_Controle", "count"),
            Total_Anomalies=("Nb_Anomalies", "sum")
        )
        .sort_values("Total_Anomalies", ascending=False)
    )

    # Synthèse par criticité
    summary_by_severity = (
        anomalies_df
        .groupby("Criticite", as_index=False)
        .agg(
            Nb_Controles=("Regle_Controle", "count"),
            Total_Anomalies=("Nb_Anomalies", "sum")
        )
        .sort_values("Total_Anomalies", ascending=False)
    )

    # Export Excel
    output_excel = os.path.join(OUTPUT_DIR, "audit_quality_report.xlsx")

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        structure_df.to_excel(writer, sheet_name="01_Structure", index=False)
        missing_report_df.to_excel(writer, sheet_name="02_Missing_Values", index=False)
        anomalies_df.to_excel(writer, sheet_name="03_Anomalies", index=False)
        summary_by_domain.to_excel(writer, sheet_name="04_Synthese_Domaine", index=False)
        summary_by_severity.to_excel(writer, sheet_name="05_Synthese_Criticite", index=False)

    # Export Markdown
    output_md = os.path.join(OUTPUT_DIR, "audit_summary.md")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write("# Synthèse audit qualité des données\n\n")
        f.write("Projet : Energy & Waste Data Transformation — Business Central, Fabric & Power BI\n\n")
        f.write(f"Date audit : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        f.write("## Synthèse par domaine\n\n")
        f.write(summary_by_domain.to_markdown(index=False))
        f.write("\n\n")

        f.write("## Synthèse par criticité\n\n")
        f.write(summary_by_severity.to_markdown(index=False))
        f.write("\n\n")

        f.write("## Top 15 contrôles avec le plus d’anomalies\n\n")
        top15 = anomalies_df.sort_values("Nb_Anomalies", ascending=False).head(15)
        f.write(top15.to_markdown(index=False))
        f.write("\n\n")

    print("\nAudit terminé avec succès.")
    print(f"Rapport Excel : {output_excel}")
    print(f"Synthèse Markdown : {output_md}")


if __name__ == "__main__":
    main()