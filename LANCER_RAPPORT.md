Analyse_HSE_Chantiers_Routiers_Maroc.pdf



# Commandes pour lancer le Rapport d'Atelier

## Démarrage des services Docker

### Lancer Watch LaTeX + Serveur PDF ensemble
```bash
docker compose -p rapport_atelier_kafka --profile dev --profile server up --no-build -d
```

Ou séparément si besoin :

#### 1. Serveur PDF uniquement (port 8082)
```bash
docker compose -p rapport_atelier_kafka --profile server up --no-build -d
```

#### 2. Watch LaTeX uniquement (compilation automatique)
```bash
docker compose -p rapport_atelier_kafka --profile dev up --no-build -d
```

## Accès au PDF généré
http://localhost:8082/main.pdf

## Arrêt des services
```bash
docker compose -p rapport_atelier_kafka down
```

## Voir les logs en temps réel

### Logs du watch (compilation LaTeX)
```bash
docker logs -f rapport-atelier-latex-watch
```
cid-logo
### Logs du serveur PDF
```bash
docker logs -f rapport-atelier-pdf-server
```

### Logs des deux services en parallèle
```bash
docker-compose -p rapport_atelier_kafka logs -f
```

### Voir uniquement les dernières 50 lignes
```bash
docker logs --tail 50 rapport-atelier-latex-watch
```

---

**Note** : Ces commandes réutilisent l'image Docker de la présentation, donc pas besoin de rebuild.
