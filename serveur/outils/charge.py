"""Simulation de charge : N utilisateurs trient chacun M contacts en même temps,
contre le serveur lancé en mode simulation (SMARTMAIL_FAUX_CLAUDE=1)."""
import concurrent.futures as cf, json, random, sys, time, urllib.request

URL = sys.argv[1]; codes = json.load(open(sys.argv[2])); M = int(sys.argv[3])
PRENOMS = ["marie", "lucas", "lea", "hugo", "chloe", "louis", "emma", "jules", "ines", "adam", "lina", "tom", "sarah", "nathan", "zoe"]
NOMS = ["durand", "martin", "petit", "roux", "morel", "fournier", "girard", "bonnet", "dupont", "lambert", "fontaine", "rousseau"]

def contacts(graine):
    r = random.Random(graine); out = []
    for i in range(M):
        p, n = r.choice(PRENOMS), r.choice(NOMS)
        k = r.random()
        if k < 0.55: out.append({"adresse": f"{p}.{n}{i}x@ecole{graine % 3}.fr", "nom": f"{p.title()} {n.title()}"})
        elif k < 0.65: out.append({"adresse": f"{p[0]}.{n}@ecole{graine % 3}.fr", "nom": f"Prof. {p.title()} {n.title()}"})
        elif k < 0.75: out.append({"adresse": f"{p}.{n}{i}@gmail.com", "nom": f"{p.title()} {n.title()}"})
        elif k < 0.80: out.append({"adresse": f"bde{i}@ecole{graine % 3}.fr", "nom": "BDE"})
        elif k < 0.90: out.append({"adresse": f"b{100000 + i}@ecole{graine % 3}.fr", "nom": ""})
        else: out.append({"adresse": f"{p}{n}{i}@ecole{graine % 3}.fr", "nom": p.title()})
    return out

def utilisateur(k):
    liste = contacts(k); t0 = time.time(); total = {"etudiant": 0, "hors": 0, "incertain": 0}; cout = 0
    for d in range(0, len(liste), 2000):
        req = urllib.request.Request(URL + "/v1/tri", data=json.dumps({"contacts": liste[d:d + 2000]}).encode(),
                                     headers={"Authorization": "Bearer " + codes[k], "Content-Type": "application/json"})
        j = json.load(urllib.request.urlopen(req, timeout=300))
        assert len(j["resultats"]) == len({c["adresse"] for c in liste[d:d + 2000]}), "résultat manquant"
        for g, n in j["resume"].items(): total[g] += n
        cout += j["cout_usd"]
    return k, time.time() - t0, total, cout

t0 = time.time()
with cf.ThreadPoolExecutor(len(codes)) as ex:
    res = list(ex.map(utilisateur, range(len(codes))))
for k, d, tot, c in res:
    print(f"utilisateur {k:2d} : {d:5.1f} s  {tot}  coût simulé {c:.3f} $")
print(f"\n{len(codes)} utilisateurs × {M} contacts en {time.time() - t0:.1f} s")
