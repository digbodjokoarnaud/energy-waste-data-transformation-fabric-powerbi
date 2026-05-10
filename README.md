# Energy & Waste Data Transformation — Business Central, Microsoft Fabric & Power BI

## Contexte

Ce projet portfolio simule une mission de transformation digitale dans le secteur énergie / environnement.

L’objectif est de remplacer progressivement des fichiers Excel métiers critiques par une architecture data gouvernée autour de :

- Microsoft Dynamics 365 Business Central
- Microsoft Fabric Lakehouse
- PySpark
- Power BI
- Power BI Service
- Gouvernance data et RLS

> Toutes les données utilisées dans ce projet sont simulées.  
> Aucune donnée réelle client n’est utilisée.

---

## Objectifs du projet

- Auditer les fichiers Excel métiers existants
- Identifier les anomalies et incohérences de données
- Nettoyer et standardiser les données
- Mapper les données vers Business Central
- Industrialiser les données dans Microsoft Fabric
- Construire une architecture Bronze / Silver / Gold
- Créer un modèle sémantique Power BI
- Concevoir des rapports Power BI métiers
- Mettre en place une gouvernance : RLS, workspaces, adoption, documentation

---

## Architecture cible

```text
Fichiers Excel métiers existants
        ↓
Audit Excel / Power Query / Python
        ↓
Mapping vers Business Central
        ↓
Business Central
        ↓
Dataflow Gen2 / API / OData
        ↓
Microsoft Fabric Lakehouse
        ↓
Bronze : données brutes
        ↓
Silver : données nettoyées avec Spark / PySpark
        ↓
Gold : tables métier prêtes pour Power BI
        ↓
Power BI Semantic Model
        ↓
Dashboards Power BI
        ↓
Power BI Service / RLS / Apps
        ↓
Formation, gouvernance et adoption