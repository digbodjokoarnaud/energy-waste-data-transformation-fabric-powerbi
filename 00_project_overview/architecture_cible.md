# Architecture cible

Cette architecture permet de passer d’un pilotage Excel manuel à une plateforme data gouvernée autour de Business Central, Microsoft Fabric et Power BI.

Business Central porte les données ERP officielles.
Fabric industrialise les traitements.
Power BI restitue les indicateurs métiers dans un cadre sécurisé

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