# Smart Mail

Trouver les étudiants d'un établissement dans l'annuaire Outlook et leur envoyer un message, un par un, depuis sa propre messagerie.

- `extension/` — extension Chrome (MV3) qui pilote Outlook web : relevé des adresses (`adresses/`), envoi (`publipostage/`), panneau à deux onglets et annuaire commun (`mega.js`).
- `serveur/` — application web installable (PWA) et API de tri des contacts par IA (FastAPI). Voir `serveur/README.md`.
- `tests/` — tests de bout en bout (Playwright, Chromium) sur de fausses pages Outlook : `publipostage.spec.ts`, `mega.spec.ts`, `app.spec.ts`, `ecoles.spec.ts`, et la campagne d'audit du pilotage d'Outlook (`audit-*.spec.ts`, `correctifs/`).

## Tests

```sh
# extension (depuis ce dossier ; node_modules = Playwright)
node node_modules/@playwright/test/cli.js test tests/publipostage.spec.ts tests/mega.spec.ts tests/correctifs --workers=3
# serveur
cd serveur && uv run pytest
```

Règles de sécurité vérifiées par les tests : jamais d'envoi à un autre destinataire que prévu, jamais deux fois, jamais de frappe ou d'effacement hors du champ visé, arrêt sans envoi sur tout état inattendu.
