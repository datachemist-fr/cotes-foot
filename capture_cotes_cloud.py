#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
capture_cotes_cloud.py  --  tourne dans le CLOUD (GitHub Actions), PC eteint OK.

A chaque passage :
  1) lit football-data (fixtures.csv + new_leagues_fixtures.csv)  -> cote Bet365 ACTUELLE
     des matchs a venir, et MEMORISE le dernier releve par match (snapshot).
     -> le dernier snapshot pris juste avant le coup d'envoi = la cote de CLOTURE.
  2) lit les fichiers resultats football-data -> score + cote de cloture officielle (B365C*)
     des matchs joues (quand dispo).
  3) met a jour snapshots_cotes.csv (la memoire) et reconstruit BASE_Cotes_Prochaine_Journee.xlsx.

Programme par .github/workflows/cotes.yml (cron, toutes les 30 min sur les heures de match).
"""
import urllib.request, csv, io, os, sys
from datetime import datetime, timedelta, timezone

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("openpyxl manquant"); sys.exit(1)

SD = os.path.dirname(os.path.abspath(__file__))
STORE = os.path.join(SD, "snapshots_cotes.csv")               # la "memoire" des cotes
OUT   = os.path.join(SD, "BASE_Cotes_Prochaine_Journee.xlsx")  # le fichier final
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

FIXTURES_MAIN  = "https://www.football-data.co.uk/fixtures.csv"
FIXTURES_WORLD = "https://www.football-data.co.uk/new_leagues_fixtures.csv"
RESULTS_MAIN   = "https://www.football-data.co.uk/mmz4281/%s/%s.csv"
RESULTS_WORLD  = "https://www.football-data.co.uk/new/%s.csv"
WINDOW_DAYS = 12   # on garde les matchs joues des 12 derniers jours + tous les a venir

LEAGUES = {
    "E0":"Angleterre Premier League","E1":"Angleterre Championship",
    "SP1":"Espagne LaLiga","SP2":"Espagne LaLiga 2",
    "I1":"Italie Serie A","I2":"Italie Serie B",
    "D1":"Allemagne Bundesliga","D2":"Allemagne Bundesliga 2",
    "F1":"France Ligue 1","F2":"France Ligue 2",
    "N1":"Pays-Bas Eredivisie","B1":"Belgique Jupiler Pro",
    "P1":"Portugal Liga","SC0":"Ecosse Premiership",
    "T1":"Turquie Super Lig","G1":"Grece Super League",
}
WORLD = {   # code /new/ : (nom lisible, Country du fixtures monde)
    "ARG":("Argentine Liga Profesional","Argentina"),"BRA":("Bresil Serie A","Brazil"),
    "MEX":("Mexique Liga MX","Mexico"),"USA":("USA MLS","USA"),
    "JPN":("Japon J1 League","Japan"),"CHN":("Chine Super League","China"),
    "DNK":("Danemark Superliga","Denmark"),"NOR":("Norvege Eliteserien","Norway"),
    "SWE":("Suede Allsvenskan","Sweden"),"SWZ":("Suisse Super League","Switzerland"),
    "AUT":("Autriche Bundesliga","Austria"),"POL":("Pologne Ekstraklasa","Poland"),
}
COUNTRY2NOM = {c:nom for nom,c in WORLD.values()}
ORDRE = list(LEAGUES.values()) + [nom for nom,_ in WORLD.values()]

FIELDS = ["champ","date","heure","dom","ext","snap_ts",
          "s1","sN","s2","sov","sun","sAHh","sAHH","sAHA",
          "bd","be","res","o1","oN","o2","oov","oun","oAHh","oAHH","oAHA","statut"]

# ---------- utilitaires ----------
def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return list(csv.DictReader(io.StringIO(r.read().decode("utf-8-sig","replace"))))

def f(row,*ks):
    for k in ks:
        v=(row.get(k) or "").strip()
        if v!="":
            try: return float(v.replace(",","."))
            except: pass
    return ""

def i(row,*ks):
    for k in ks:
        v=(row.get(k) or "").strip()
        if v!="":
            try: return int(float(v))
            except: pass
    return ""

def dt(d,t=""):
    for fmt in ("%d/%m/%Y %H:%M","%d/%m/%y %H:%M","%d/%m/%Y","%d/%m/%y"):
        try: return datetime.strptime((d+" "+t).strip(),fmt)
        except: pass
    return None

def key(champ,d,dom,ext): return (champ,d,dom.strip().lower(),ext.strip().lower())

def season_code():
    now=datetime.now(timezone.utc)
    y=now.year if now.month>=7 else now.year-1
    return "%02d%02d"%(y%100,(y+1)%100)

# ---------- charge / sauvegarde la memoire ----------
def load_store():
    st={}
    if os.path.exists(STORE):
        with open(STORE,encoding="utf-8") as fp:
            for row in csv.DictReader(fp,delimiter=";"):
                st[key(row["champ"],row["date"],row["dom"],row["ext"])]=row
    return st

def save_store(st):
    rows=sorted(st.values(), key=lambda r:(dt(r["date"],r.get("heure","")) or datetime.max, r["champ"]))
    with open(STORE,"w",newline="",encoding="utf-8") as fp:
        w=csv.DictWriter(fp,fieldnames=FIELDS,delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k:r.get(k,"") for k in FIELDS})

def blank(champ,d,h,dom,ext):
    r={k:"" for k in FIELDS}
    r.update(champ=champ,date=d,heure=h,dom=dom,ext=ext,statut="A venir (en cours)")
    return r

# ---------- mise a jour depuis football-data ----------
def maj(st):
    now=datetime.now(timezone.utc); ts=now.strftime("%Y-%m-%d %H:%M")
    limite=now.replace(tzinfo=None)-timedelta(days=WINDOW_DAYS)
    # 1) matchs a venir -> snapshot de la cote actuelle (= futur "closing")
    try:
        for r in fetch(FIXTURES_MAIN):
            div=(r.get("Div") or "").strip()
            if div not in LEAGUES: continue
            champ=LEAGUES[div]; d=(r.get("Date") or "").strip()
            dom=(r.get("HomeTeam") or "").strip(); ext=(r.get("AwayTeam") or "").strip()
            if not(dom and ext and d): continue
            k=key(champ,d,dom,ext)
            row=st.get(k) or blank(champ,d,(r.get("Time") or "").strip(),dom,ext)
            if row.get("statut","").startswith("Joue"): continue  # deja joue, on ne touche plus
            row.update(heure=(r.get("Time") or "").strip(), snap_ts=ts,
                       s1=f(r,"B365H"),sN=f(r,"B365D"),s2=f(r,"B365A"),
                       sov=f(r,"B365>2.5"),sun=f(r,"B365<2.5"),
                       sAHh=f(r,"AHh"),sAHH=f(r,"B365AHH"),sAHA=f(r,"B365AHA"))
            st[k]=row
    except Exception as e: print("fixtures.csv:",e)
    try:
        for r in fetch(FIXTURES_WORLD):
            pays=(r.get("Country") or "").strip()
            if pays not in COUNTRY2NOM: continue
            champ=COUNTRY2NOM[pays]; d=(r.get("Date") or "").strip()
            dom=(r.get("Home") or "").strip(); ext=(r.get("Away") or "").strip()
            if not(dom and ext and d): continue
            k=key(champ,d,dom,ext)
            row=st.get(k) or blank(champ,d,(r.get("Time") or "").strip(),dom,ext)
            if row.get("statut","").startswith("Joue"): continue
            row.update(heure=(r.get("Time") or "").strip(), snap_ts=ts,
                       s1=f(r,"B365H"),sN=f(r,"B365D"),s2=f(r,"B365A"))
            st[k]=row
    except Exception as e: print("new_leagues_fixtures.csv:",e)
    # 2) matchs joues -> score + cote cloture officielle (B365C*) si dispo
    season=season_code()
    def ingest(rows,champ):
        for r in rows:
            dom=(r.get("HomeTeam") or r.get("Home") or "").strip()
            ext=(r.get("AwayTeam") or r.get("Away") or "").strip()
            d=(r.get("Date") or "").strip()
            if not(dom and ext and d): continue
            dd=dt(d)
            if dd is None or dd<limite or dd>now.replace(tzinfo=None)+timedelta(days=1): continue
            bd,be=i(r,"FTHG"),i(r,"FTAG")
            if bd=="" or be=="": continue
            res="1" if bd>be else ("2" if bd<be else "N")
            k=key(champ,d,dom,ext)
            row=st.get(k) or blank(champ,d,(r.get("Time") or "").strip(),dom,ext)
            row.update(statut="Joue (cloture)",bd=bd,be=be,res=res,
                       o1=f(r,"B365CH"),oN=f(r,"B365CD"),o2=f(r,"B365CA"),
                       oov=f(r,"B365C>2.5"),oun=f(r,"B365C<2.5"),
                       oAHh=f(r,"AHCh"),oAHH=f(r,"B365CAHH"),oAHA=f(r,"B365CAHA"))
            # si pas de snapshot pre-match (ex. 1er passage), on prend la cote pre-match du fichier resultats
            if row.get("s1","")=="":
                row.update(s1=f(r,"B365H"),sN=f(r,"B365D"),s2=f(r,"B365A"),
                           sov=f(r,"B365>2.5"),sun=f(r,"B365<2.5"),
                           sAHh=f(r,"AHh"),sAHH=f(r,"B365AHH"),sAHA=f(r,"B365AHA"))
            st[k]=row
    for code,champ in LEAGUES.items():
        try: ingest(fetch(RESULTS_MAIN%(season,code)),champ)
        except Exception as e: print("res",code,e)
    for code,(champ,_) in WORLD.items():
        try: ingest(fetch(RESULTS_WORLD%code),champ)
        except Exception as e: print("res",code,e)
    # 3) purge des vieux matchs
    for k in list(st):
        dd=dt(st[k]["date"],st[k].get("heure",""))
        if dd and dd<limite: del st[k]

# ---------- construction de l'Excel ----------
GREEN=PatternFill("solid",fgColor="C6EFCE"); JAUNE=PatternFill("solid",fgColor="FFEB9C")
ROSE=PatternFill("solid",fgColor="FFC7CE"); BLEU=PatternFill("solid",fgColor="DDEBF7")
HF=Font(bold=True,color="FFFFFF"); HEADER=PatternFill("solid",fgColor="1F4E78")
THIN=Side(style="thin",color="D9D9D9"); BORDER=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
CENTER=Alignment(horizontal="center",vertical="center")
HEADS=["Date","Heure","Championnat","Statut","Domicile","Exterieur","Buts Dom","Buts Ext",
       "Resultat","Issue favori","Cote 1 C","Cote N C","Cote 2 C","Over 2.5 C","Under 2.5 C"]
WIDTHS=[11,7,24,18,18,18,8,8,8,12,8,8,8,10,10]

def num(x):
    try: return float(x)
    except: return None

def build(st):
    wb=openpyxl.Workbook(); glob=wb.active; glob.title="Cotes journee"
    def entete(ws):
        ws.append(HEADS)
        for c in range(1,len(HEADS)+1):
            cell=ws.cell(1,c); cell.font=HF; cell.fill=HEADER; cell.alignment=CENTER; cell.border=BORDER
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width=WIDTHS[c-1]
        ws.freeze_panes="A2"
    def ligne(ws,r):
        joue=r.get("statut","").startswith("Joue")
        # cote cloture = officielle si dispo, sinon dernier snapshot pre-match
        def pick(o,s):
            v=num(r.get(o,""));
            return v if v is not None else num(r.get(s,""))
        o1=pick("o1","s1"); oN=pick("oN","sN"); o2=pick("o2","s2")
        oov=pick("oov","sov"); oun=pick("oun","sun")
        res=r.get("res","") or ""
        fav=""
        trio=[("1",o1),("N",oN),("2",o2)]; nn=[(k,v) for k,v in trio if v is not None]
        if res and len(nn)>=2:
            favk=min(nn,key=lambda x:x[1])[0]; fav="Favori" if res==favk else "Autre"
        bd=r.get("bd",""); be=r.get("be","")
        ws.append([r.get("date",""),r.get("heure",""),r.get("champ",""),r.get("statut",""),
                   r.get("dom",""),r.get("ext",""),
                   int(bd) if str(bd).isdigit() else "", int(be) if str(be).isdigit() else "",
                   res,fav,o1,oN,o2,oov,oun])
        rr=ws.max_row
        for c in range(1,len(HEADS)+1):
            cell=ws.cell(rr,c); cell.border=BORDER
            if c>=7: cell.alignment=CENTER
        if not joue: ws.cell(rr,4).fill=BLEU
        if fav=="Favori": ws.cell(rr,10).fill=GREEN
        elif fav=="Autre": ws.cell(rr,10).fill=ROSE
        COLR={"1":(11,GREEN),"N":(12,JAUNE),"2":(13,ROSE)}
        if res in COLR:
            col,fill=COLR[res]; ws.cell(rr,col).fill=fill; ws.cell(rr,9).fill=fill
        # Over/Under : colorie la cote gagnante selon le total de buts (vert=favori, rose=outsider)
        if str(bd).lstrip("-").isdigit() and str(be).lstrip("-").isdigit() and oov is not None and oun is not None:
            tot=int(bd)+int(be); wc=14 if tot>=3 else 15
            wcote,lcote=(oov,oun) if wc==14 else (oun,oov)
            ws.cell(rr,wc).fill=GREEN if wcote<=lcote else ROSE
    entete(glob)
    rows=sorted(st.values(),key=lambda r:(dt(r["date"],r.get("heure","")) or datetime.max,r["champ"]))
    for r in rows: ligne(glob,r)
    par={}
    for r in rows: par.setdefault(r["champ"],[]).append(r)
    for champ in ORDRE:
        if champ in par:
            ws=wb.create_sheet(champ[:31]); entete(ws)
            for r in par[champ]: ligne(ws,r)
    wb.save(OUT)

def main():
    st=load_store()
    maj(st)
    save_store(st)
    build(st)
    av=sum(1 for r in st.values() if r.get("statut","").startswith("A venir"))
    jo=len(st)-av
    print("OK : %d matchs (%d a venir, %d joues) | %s"%(len(st),av,jo,
          datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")))

if __name__=="__main__":
    main()
