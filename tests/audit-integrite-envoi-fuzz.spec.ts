// Audit « Intégrité de chaque envoi » : exploration aléatoire à graines.
// Pour chaque graine, la fausse page « nouvel Outlook » (tests/audit/integrite-envoi-fuzz/)
// varie (délais de résolution, suggestion voisine qui remplace l'adresse, pastille
// au nom seul, résolution qui échoue, doublons, collage refusé ou filtré, restes
// d'un message précédent, signature avant/après/au milieu, doublons cachés,
// enveloppes, identifiants changeants, re-rendus) et le plan d'envoi aussi
// (adresses variées, Cci par paquets de 1 à 50, importance, signature retirée).
// Ce que la fausse page a réellement « envoyé » est comparé au prévu : identique,
// ou aucun envoi et un message clair. Toute divergence fait échouer la graine.
//   node node_modules/@playwright/test/cli.js test tests/audit-integrite-envoi-fuzz.spec.ts --workers=2 --reporter=line
//
// Graines qui échouent aujourd'hui (générateur figé, 220 graines) et cause dans le code ;
// reproductions minimales : audit-integrite-envoi-fuzz-cibles.spec.ts.
// BUG : accents dans la partie locale coupés vers une AUTRE adresse (« zoë.nguyen@ » →
//       « nguyen@ », « andré.martin@ » → « martin@ »), « josé@ » ignoré sans message —
//       moteur.js:60-69 (extraireAdresses, ASCII seul). Graines 24 32 39 43 44 53 58 74 84 103 180 214.
// BUG : pastille jamais comparée à l'adresse collée (suggestion voisine visible acceptée) —
//       outlook.js:367-372, 396-406, 452-458. Graines 24 105 133.
// BUG : aucun contrôle final avant « Envoyer » (corps réécrit par un complément, pastille
//       doublée ou remplacée à la sortie du champ, importance perdue, signature reposée) —
//       moteur.js:304-313, outlook.js:719-728. Graines 27 75 82 88 92 109 210 220.
// BUG : corps vérifié sur 30 caractères + présence des liens n'importe où (signature
//       comprise) : corps tronqué, lien perdu ou réécrit masqué — outlook.js:558-565.
//       Graines 3 12 25 35 49 50 68 71 72 81 107 120 202.
// BUG : texte déjà présent dans un corps neuf pris pour la signature (reste envoyé) —
//       outlook.js:551-552. Graines 22 46 183.
// BUG : Cc jamais contrôlé, Cci non contrôlé en mode « À » — outlook.js:386-409, 471.
//       Graines 24 84 105 147 160 208.
// BUG : rédaction trop profonde (> 25 niveaux) : clic sur le premier « Envoyer » de la page
//       (conversation Teams) — outlook.js:154-163, 723. Graines 18 136 156.
// BUG : visible() accepte hors écran / opacité 0 / hauteur nulle : saisie et envoi dans une
//       rédaction masquée — outlook.js:67-75, 146-150. Graines 67 151.
// BUG : pastille invalide comptée comme validée ; refus explicite d'Outlook classé « à
//       vérifier » et rangé dans « Déjà envoyés » — outlook.js:367-372, moteur.js:316, 399.
//       Graines 155 214.
import { test, expect, executer, rapport } from "./audit/integrite-envoi-fuzz/banc";
import { genererGraine, verifier } from "./audit/integrite-envoi-fuzz/generateur";

test.describe.configure({ mode: "parallel" });

const N = Number(process.env.FUZZ_N || 220);
const DEBUT = Number(process.env.FUZZ_DEBUT || 1);

for (let g = DEBUT; g < DEBUT + N; g++) {
  const { cfg, plan } = genererGraine(g);
  test(`graine ${g} [${plan.resume.join(", ")}]`, async ({ banc }) => {
    const r = await executer(banc, cfg, plan);
    const d = verifier(plan, r);
    if (process.env.FUZZ_TRACE) console.log(rapport(plan, r, d));
    expect(d, rapport(plan, r, d)).toEqual([]);
  });
}
