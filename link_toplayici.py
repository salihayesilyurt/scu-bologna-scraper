# -*- coding: utf-8 -*-
"""
Bologna Program Sayfasi -> links.txt  (OTOMATIK link toplayici)
================================================================
Bir program (mufredat) sayfasindaki TUM derslerin detay linklerini
otomatik toplar ve links.txt olarak kaydeder. Boylece her ders icin
elle 'i' butonuna tiklayip URL kopyalamaniza gerek kalmaz.

NASIL CALISIR
-------------
progCourses.aspx ders listesini verir ama her dersin 'i' (detay) butonu
ASP.NET '__doPostBack' ile calisir; curCourse ID'si sayfada yazmaz.
Bu arac, tarayicinin yaptigini yapar: her 'i' butonu icin sunucuya bir
'postback' gonderir, donen cevaptaki progCourseDetails.aspx?curCourse=...
adresini yakalar.

KULLANIM
--------
    pip install requests beautifulsoup4

    python link_toplayici.py "https://obs.cumhuriyet.edu.tr/oibs/bologna/index.aspx?lang=tr&curOp=showPac&curUnit=33&curSunit=123"

veya sadece curSunit numarasiyla:

    python link_toplayici.py 123

Sonra normal sekilde:  python bologna_scraper.py
"""

import argparse
import re
import sys
import time

import requests
from bs4 import BeautifulSoup

BASE      = "https://obs.cumhuriyet.edu.tr/oibs/bologna/"
LISTE_URL = BASE + "progCourses.aspx?curSunit={sunit}&lang=tr"
DETAY_URL = BASE + "progCourseDetails.aspx?curCourse={id}&lang=tr"
HEADERS   = {"User-Agent": "Mozilla/5.0 (link-collector)"}
BEKLE     = 0.4   # istekler arasi saniye

# 'i' (ders detay) butonunun postback hedef kalibi
HEDEF_KALIBI = r"grdBolognaDersler\$ctl\d+\$btnDersAyrinti"


def curSunit_bul(arg):
    """Verilen metinden curSunit numarasini cikarir (URL ya da duz sayi)."""
    m = re.search(r"curSunit=(\d+)", arg or "")
    if m:
        return m.group(1)
    a = (arg or "").strip()
    return a if a.isdigit() else None


def form_alanlari(soup):
    """Sayfadaki <form> icindeki gonderilebilir tum alanlari toplar
    (__VIEWSTATE, __EVENTVALIDATION, gizli alanlar, select degerleri...)."""
    alanlar = {}
    kapsam = soup.find("form") or soup
    for inp in kapsam.find_all("input"):
        ad = inp.get("name")
        if not ad:
            continue
        tip = (inp.get("type") or "text").lower()
        if tip in ("submit", "button", "image", "reset"):
            continue
        if tip in ("checkbox", "radio"):
            if inp.has_attr("checked"):
                alanlar[ad] = inp.get("value", "on")
        else:
            alanlar[ad] = inp.get("value", "")
    for sel in kapsam.find_all("select"):
        ad = sel.get("name")
        if not ad:
            continue
        opt = sel.find("option", selected=True) or sel.find("option")
        if opt is not None:
            alanlar[ad] = opt.get("value", opt.get_text(strip=True))
    return alanlar


def detay_hedefleri(html):
    """progCourses sayfasindaki tum 'i' (ders detay) postback hedeflerini,
    sayfadaki sirayla ve tekrarsiz dondurur."""
    gor, sirali = set(), []
    for h in re.findall(HEDEF_KALIBI, html):
        if h not in gor:
            gor.add(h)
            sirali.append(h)
    return sirali


def id_bul(metin):
    """Bir metin ya da URL icinden curCourse ID'sini cikarir."""
    m = re.search(r"curCourse=(\d+)", metin or "")
    return m.group(1) if m else None


def collect(sunit, bekle=BEKLE):
    """curSunit icin tum ders curCourse ID'lerini toplar.
    Geri donus: (benzersiz_id_listesi, hatali_hedef_listesi)."""
    s = requests.Session()
    s.headers.update(HEADERS)
    liste_url = LISTE_URL.format(sunit=sunit)

    r = s.get(liste_url, timeout=30)
    r.raise_for_status()
    taban_alan = form_alanlari(BeautifulSoup(r.text, "html.parser"))
    hedefler = detay_hedefleri(r.text)
    if not hedefler:
        sys.exit("HATA: Sayfada ders bulunamadi. curSunit dogru mu, "
                 "yoksa site yapisi mi degisti?")
    print(f"{len(hedefler)} ders satiri bulundu. curCourse ID'leri toplaniyor...\n")

    def postla(alanlar, hedef):
        veri = dict(alanlar)
        veri["__EVENTTARGET"]   = hedef
        veri["__EVENTARGUMENT"] = ""
        return s.post(liste_url, data=veri, timeout=30,
                      headers={"Referer": liste_url}, allow_redirects=True)

    idler, hatali = [], []
    for i, hedef in enumerate(hedefler, 1):
        try:
            pr  = postla(taban_alan, hedef)
            cid = id_bul(pr.url) or id_bul(pr.text)
            if not cid:
                # __VIEWSTATE eskimis olabilir: sayfayi yenile ve tekrar dene
                r2 = s.get(liste_url, timeout=30)
                yeni = form_alanlari(BeautifulSoup(r2.text, "html.parser"))
                pr  = postla(yeni, hedef)
                cid = id_bul(pr.url) or id_bul(pr.text)
            if cid:
                idler.append(cid)
                print(f"  [{i}/{len(hedefler)}] curCourse={cid}")
            else:
                hatali.append(hedef)
                print(f"  [{i}/{len(hedefler)}] {hedef} -> ID alinamadi")
        except Exception as e:
            hatali.append(hedef)
            print(f"  [{i}/{len(hedefler)}] {hedef} -> HATA: {e}")
        time.sleep(bekle)

    gor, benzersiz = set(), []           # tekrarsiz, sira korunarak
    for x in idler:
        if x not in gor:
            gor.add(x)
            benzersiz.append(x)
    return benzersiz, hatali


def main():
    ap = argparse.ArgumentParser(description="Bologna program sayfasi -> links.txt")
    ap.add_argument("program", help="program sayfasi URL'si VEYA curSunit numarasi")
    ap.add_argument("--out", default="links.txt", help="cikti dosyasi (varsayilan links.txt)")
    ap.add_argument("--bekle", type=float, default=BEKLE, help="istekler arasi saniye")
    args = ap.parse_args()

    sunit = curSunit_bul(args.program)
    if not sunit:
        sys.exit("HATA: curSunit bulunamadi. Program linkini ya da numarasini verin.\n"
                 "Ornek: python link_toplayici.py 123")

    idler, hatali = collect(sunit, args.bekle)
    if not idler:
        sys.exit("\nHic curCourse ID toplanamadi; links.txt yazilmadi.")

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(f"# Otomatik olusturuldu  |  curSunit={sunit}  |  {len(idler)} ders\n")
        for cid in idler:
            f.write(DETAY_URL.format(id=cid) + "\n")

    print(f"\nBitti. {len(idler)} ders linki yazildi -> {args.out}")
    if hatali:
        print(f"{len(hatali)} satir icin ID alinamadi (atlandi). "
              "Tekrar calistirmayi deneyebilirsiniz.")
    print("\nSimdi su komutu calistirin:  python bologna_scraper.py")


if __name__ == "__main__":
    main()
