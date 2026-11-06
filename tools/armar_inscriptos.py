#!/usr/bin/env python3
"""Arma assets/inscriptos.csv y assets/escuelas-actualizadas.csv a partir del PDF de SET-ONLINE (sportdata).
Uso: python3 tools/leer_pdf_sportdata.py inscriptos.pdf inscriptos.json
     python3 tools/armar_inscriptos.py planilla.xlsx inscriptos.json r1
Reglas de escuela: tools/reglas_escuelas.json (por club del sistema y por competidor)."""
import json, re, collections, sys, csv
from pathlib import Path
HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
import build
from build import nrm, close

xlsx, js, ev = sys.argv[1], sys.argv[2], sys.argv[3]
REG = json.load(open(HERE / "tools/reglas_escuelas.json", encoding="utf-8"))
d = json.load(open(js, encoding="utf-8"))
cats = build.parse(xlsx)
build.actualizar_escuelas.__globals__  # noqa
TEAM = {"53570": "Pre-Juveniles", "54570": "Juveniles", "55580": "Adultos"}
key = lambda c: ("LPE-" + c["age"]) if c["team"] else c["num"]
bykey = {key(c): c for c in cats}

def clean(x):
    x = re.sub(r"\b(EQUIPO|INSCRIPCIONES|CATEGORÍA)\b", "", x)
    x = re.sub(r"ARGENTINA \(ARG\).*$", "", x)
    return re.sub(r"\s+", " ", x).strip().rstrip(",")
def club(s):
    m = re.findall(r"\(([^()]*)\)", s)
    return (m[-1] if m else s).strip().strip(".").upper()
def parts(n):
    sur, _, giv = n.partition(","); return nrm(sur).split(), nrm(giv).split()
def match(rowname, T, fuzzy):
    S, G = parts(rowname)
    if not S or not G: return False
    rest = [t for t in T if t not in S]
    if all(s in T for s in S) and G[0] in rest: return True
    if S[0] in T and G[0] in T: return True
    return fuzzy and all(s in T for s in S) and any(close(G[0], t) for t in rest)

persons = [(c, r) for c in cats if not c["team"] for r in c["rows"]]
given = {g for c, r in persons for g in parts(r["n"])[1]}
FIX = REG.get("nombres", {})
def fmt_name(raw):
    if raw in FIX: return FIX[raw]
    T = raw.split(); tt = lambda w: w[:1].upper() + w[1:].lower()
    if len(T) <= 2: return tt(T[0]) + (", " + tt(T[1]) if len(T) == 2 else "")
    i = next((k for k in range(1, len(T)) if nrm(T[k]).strip() in given), 1)
    return " ".join(map(tt, T[:i])) + ", " + " ".join(map(tt, T[i:]))

recs = []
for x in d:
    if x["cod"] == "00000": continue
    x["name"], x["school"] = clean(x["name"]), clean(x["school"]); x["club"] = club(x["school"])
    c = bykey.get("LPE-" + TEAM[x["cod"]] if x["cod"] in TEAM else x["cod"])
    if not c: print("SIN CATEGORIA", x); continue
    recs.append({"x": x, "c": c, "T": nrm(x["name"]).split()})

# escuela dominante de cada club (según inscriptos ya conocidos)
cm = collections.defaultdict(collections.Counter)
for R in recs:
    if R["c"]["team"]: continue
    for fz in (False, True):
        hits = [r for c, r in persons if match(r["n"], R["T"], fz)]
        if hits: break
    for r in hits: cm[R["x"]["club"]][r["s"]] += 1
dom = {k: v.most_common(1)[0][0] for k, v in cm.items()}
dom.update(REG.get("club_escuela_nuevos", {}))

def resolve(R):
    """Devuelve (nombre_en_ranking|None, escuela_actual|None)."""
    T, c, cl = R["T"], R["c"], R["x"]["club"]
    for fz in (False, True):
        inc = [r for r in c["rows"] if match(r["n"], T, fz)]
        glob = [r for cc, r in persons if match(r["n"], T, fz)]
        if inc or glob: break
    pool = inc or glob
    if not pool: return None, None
    # si hay homónimos, preferir el de la escuela del club
    pref = [r for r in pool if r["s"] == dom.get(cl)]
    if inc and not pref:
        alt = [r for r in glob if r["s"] == dom.get(cl) and r["n"] == inc[0]["n"]]
        if alt: pref = alt
    r = (pref or pool)[0]
    return r["n"], r["s"]

def final_school(name, cur, cl):
    per = REG.get("competidor", {})
    if name and name in per:
        v = per[name]
        return cur if v == "mantener" else (dom.get(cl, cur) if v == "club" else v)
    rule = REG.get("club", {}).get(cl)
    if rule == "mantener" or (rule is None and cur): return cur or dom.get(cl, cl.title())
    if rule and "{escuela}" in rule:
        return rule.format(escuela=cur) if cur and cur not in REG.get("sin_sufijo", []) else REG.get("club_base", {}).get(cl, cl)
    if rule: return rule
    return dom.get(cl, cl.title())

def completar(n, raw):
    """Si el nombre del PDF es más completo que el del ranking, devuelve la versión completa."""
    S, G = parts(n); T = raw.split(); TN = [nrm(t).strip() for t in T]
    if not S or not G or TN[0] != S[0] or len(T) <= len(S) + len(G): return n
    k = next((i for i in range(1, len(TN)) if TN[i] == G[0]), None)
    if k is None or not all(x in TN[k:] for x in G) or not all(x in TN[:k] for x in S): return n
    tt = lambda w: "-".join(p[:1].upper() + p[1:].lower() for p in w.split("-"))
    return " ".join(map(tt, T[:k])) + ", " + " ".join(map(tt, T[k:]))

out, upd, seen_upd = [], [], set()
ren, seen_ren = [], set()
indiv_by_club = collections.defaultdict(set)
for R in recs:
    if R["c"]["team"]: continue
    n, s = resolve(R)
    if n:
        full = completar(n, R["x"]["name"])
        if full != n and n not in seen_ren:
            seen_ren.add(n); ren.append({"nombre_anterior": n, "escuela": s or "", "nombre_nuevo": full})
        n = full
    R["n"], R["cur"] = (n or fmt_name(R["x"]["name"])), s
    R["s"] = final_school(n, s, R["x"]["club"])
    indiv_by_club[R["x"]["club"]].add(R["n"])
    if n and s and R["s"] != s and (n, s) not in seen_upd:
        seen_upd.add((n, s)); upd.append({"nombre": n, "escuela_anterior": s, "escuela_nueva": R["s"]})
    out.append({"evento": ev, "codigo": key(R["c"]), "nombre": R["n"], "escuela": R["s"], "integrantes": ""})
# competidores que solo participan como integrantes de un equipo (código 00000): sirven para completar nombres
for x in d:
    if x["cod"] == "00000":
        indiv_by_club[club(clean(x["school"]))].add(fmt_name(clean(x["name"])))
for R in recs:
    if not R["c"]["team"]: continue
    c, x = R["c"], R["x"]
    mt = re.match(r"^(.*?)\((.*?)\)?\s*$", x["name"])
    tn = (mt.group(1) if mt else x["name"]).strip(); mem = [m.strip(" )") for m in (mt.group(2).split(",") if mt else [])]
    hit = None
    for r in c["rows"]:
        rm = nrm(" ".join(r.get("m", [])).replace(",", " ")).split()
        if mem and all(nrm(m).split()[0] in rm for m in mem): hit = r; break
    if hit:
        out.append({"evento": ev, "codigo": key(c), "nombre": hit["n"], "escuela": hit["s"], "integrantes": " / ".join(hit.get("m", []))}); continue
    full = REG.get("equipos", {}).get(tn, [])
    mems = []
    for k, m in enumerate(mem):
        if k < len(full): mems.append(full[k]); continue
        S = nrm(m).split()
        cand = sorted(n for n in indiv_by_club[x["club"]] if nrm(n.replace(",", " ")).split()[:len(S)] == S)
        full_m = REG.get("integrantes", {}).get(nrm(m).strip().upper()) or REG.get("integrantes", {}).get(m.strip().upper())
        mems.append(full_m or (cand[0] if len(cand) == 1 else m.title()))
    tname = REG.get("nombres_equipos", {}).get(tn, " ".join(w if w.isupper() and len(w) <= 4 else w.capitalize() for w in tn.split()))
    out.append({"evento": ev, "codigo": key(c), "nombre": tname, "escuela": final_school(None, None, x["club"]), "integrantes": " / ".join(mems)})
# aplicar los nombres completados a todas las filas (un mismo competidor puede figurar más corto en otra categoría)
rmap = {r["nombre_anterior"]: r["nombre_nuevo"] for r in ren}
for o in out:
    if o["nombre"] in rmap: o["nombre"] = rmap[o["nombre"]]
    if o["integrantes"]: o["integrantes"] = " / ".join(rmap.get(x.strip(), x.strip()) for x in o["integrantes"].split("/"))
renom = REG.get("renombrar_escuelas", {})
for o in out:
    o["escuela"] = renom.get(o["escuela"], o["escuela"])
for u in upd:
    u["escuela_nueva"] = renom.get(u["escuela_nueva"], u["escuela_nueva"])
seen = {o["codigo"] for o in out}
for c in cats:
    if key(c) not in seen: out.append({"evento": ev, "codigo": key(c), "nombre": "-", "escuela": "", "integrantes": ""})
with open(HERE / "assets/inscriptos.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["evento", "codigo", "nombre", "escuela", "integrantes"], delimiter=";"); w.writeheader(); w.writerows(out)
for ant, nue in REG.get("renombrar_escuelas", {}).items():
    upd.append({"nombre": "*", "escuela_anterior": ant, "escuela_nueva": nue})
with open(HERE / "assets/nombres-actualizados.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["nombre_anterior", "escuela", "nombre_nuevo"], delimiter=";"); w.writeheader(); w.writerows(ren)
print(f"{len(ren)} nombres completados desde el PDF")
for r in ren: print("  ", r["nombre_anterior"], "->", r["nombre_nuevo"])
with open(HERE / "assets/escuelas-actualizadas.csv", "w", encoding="utf-8", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["nombre", "escuela_anterior", "escuela_nueva"], delimiter=";"); w.writeheader(); w.writerows(upd)
print(f"{len(out)} inscripciones · {len(upd)} cambios de escuela")
for u in upd: print("  ", u["nombre"], ":", u["escuela_anterior"], "->", u["escuela_nueva"])
