# 05 — Microsoft Fabric Lakehouse

Cette étape met en place une architecture Lakehouse avec une approche médaillon.

## Architecture

- Bronze : données brutes
- Silver : données nettoyées
- Gold : tables métier pour Power BI

## Transformations

Les transformations sont réalisées avec PySpark dans des notebooks Microsoft Fabric.

## Tables Gold principales

- gold_dim_client
- gold_dim_site
- gold_fact_tonnage
- gold_fact_valorisation_energie
- gold_fact_facturation
- gold_fact_maintenance
- gold_quality_checks