# 06 — Modèle sémantique Power BI

Le modèle Power BI repose sur une architecture en étoile.

## Dimensions

- DimDate
- DimClient
- DimSite
- DimContrat
- DimTypeDechet
- DimPrestation

## Tables de faits

- FactTonnage
- FactValorisationEnergie
- FactFacturation
- FactMaintenance
- FactMigrationExcel
- FactQualityChecks

## Objectif

Centraliser les mesures DAX et les règles métier dans un modèle sémantique fiable et gouverné.