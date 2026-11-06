# Smart Mail — serveur et application

L'application (PWA installable sur téléphone et ordinateur, sans App Store ni réglage de sécurité) et l'API de tri des contacts par IA.

## Essayer tout de suite (sur ce Mac)

Double-clic sur **`Smart Mail - démo.command`** (dossier parent) : faux annuaire d'école, faux Claude, aucun mail envoyé.
Ou à la main :

```sh
SMARTMAIL_DEMO=1 SMARTMAIL_FAUX_CLAUDE=1 SMARTMAIL_FAUX_LATENCE=1 uv run uvicorn app.main:app --port 8766
```

## Comment ça trie

1. **Règles** (gratuit) : messageries perso, boîtes génériques, associations, numéros étudiants, sous-domaines étudiants/personnel.
2. **Logique de l'école** (`POST /v1/logique`) : Claude lit un échantillon d'environ 30 personnes et en déduit les formats d'adresse des étudiants et du personnel ; le serveur l'applique ensuite lui-même (gratuit).
3. **Claude** (`POST /v1/tri`) pour ce qui reste, par lots de 150, réponse JSON imposée. Sans indice, la réponse est « incertain », jamais une devinette.
4. **Cache** 180 jours : une adresse déjà triée ne coûte plus rien.

Garde-fous : code d'accès et quota mensuel par personne, plafond de dépense global (`SMARTMAIL_BUDGET_MENSUEL_USD`), 8 appels simultanés au plus à Anthropic.

## Mettre en ligne sur le VPS (Ubuntu 24.04)

1. Sur le VPS : installer Docker (`curl -fsSL https://get.docker.com | sh`).
2. Ici : `deploiement/deploy.sh ubuntu@IP-DU-VPS` (copie le code).
3. Sur le VPS : `cp smartmail/deploiement/.env.exemple smartmail/deploiement/.env`, puis remplir `DOMAINE` (ex. `51-83-12-34.sslip.io` = l'IP avec des tirets : HTTPS gratuit et automatique) et `ANTHROPIC_API_KEY`.
4. Relancer `deploiement/deploy.sh ubuntu@IP-DU-VPS`. L'app est sur `https://DOMAINE`.

## Administration (sur le VPS)

```sh
docker compose exec smartmail python -m app.admin creer "Pierre" --quota 20000   # affiche le code UNE fois
docker compose exec smartmail python -m app.admin liste                          # consommation et dépense du mois
docker compose exec smartmail python -m app.admin quota "Pierre" 50000
docker compose exec smartmail python -m app.admin desactiver "Pierre"
```

## Tests

```sh
uv run pytest                                        # 141 tests serveur (Claude simulé), dont le banc d'essai des 8 écoles
uv run python -m outils.eval_ecoles --contacts 600   # estimation ; ajouter --oui pour évaluer avec le VRAI Claude (~5 $)
uv run python outils/charge.py URL codes.json 3000   # charge : N utilisateurs en parallèle
# Application (depuis le dossier parent, serveur démo lancé sur 8766) :
APP_URL=http://127.0.0.1:8766 node node_modules/@playwright/test/cli.js test tests/app.spec.ts tests/ecoles.spec.ts
```

Banc d'essai (`app/ecoles.py`) : 8 écoles fictives aux conventions différentes (prenom.nom / p.nom, numéros étudiants,
sous-domaine etu., intitulés seuls, nom.prenom, multi-campus, anglophone, annuaire « sale » avec salles, listes,
invités #EXT#, partenaires, contacts perso, noms composés). Règle d'or vérifiée partout : aucun non-étudiant retenu.
Démo d'une école : `http://127.0.0.1:8766/?ecole=numeros`.

## Reste à faire

- Connexion Microsoft réelle (OAuth) + lecture de l'annuaire par Microsoft Graph, à la place de `/demo/annuaire`.
- Envoi réel en arrière-plan par Microsoft Graph (aujourd'hui l'écran d'envoi est une démonstration).
