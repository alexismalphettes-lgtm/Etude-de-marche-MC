"""Script de diagnostic - affiche la structure HTML d'une page occasion AN66."""
import time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

URL = "https://www.an66.fr/occasions?id_annonce=4767"

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36",
        locale="fr-FR",
    )
    page = context.new_page()
    page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page.goto(URL, wait_until="load", timeout=60000)
    time.sleep(4)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "html.parser")

print("=" * 60)
print("TITRES H1 / H2 / H3 :")
print("=" * 60)
for tag in soup.find_all(["h1", "h2", "h3"]):
    print(f"  <{tag.name}> : {tag.get_text(strip=True)[:100]}")

print()
print("=" * 60)
print("TEXTE CONTENANT 'prix' :")
print("=" * 60)
for el in soup.find_all(True):
    t = el.get_text(strip=True)
    if "prix" in t.lower() and "€" in t and len(t) < 80:
        print(f"  <{el.name} class='{el.get('class','')}'>  {t}")

print()
print("=" * 60)
print("TEXTE CONTENANT 'moteur' ou 'motoris' :")
print("=" * 60)
import re
for el in soup.find_all(True):
    t = el.get_text(separator=" ", strip=True)
    if re.search(r"moteur|motoris", t, re.I) and len(t) < 200:
        print(f"  <{el.name} class='{el.get('class','')}'>  {t[:150]}")

print()
print("=" * 60)
print("TEXTE CONTENANT 'heure' :")
print("=" * 60)
for el in soup.find_all(True):
    t = el.get_text(separator=" ", strip=True)
    if re.search(r"heure", t, re.I) and len(t) < 200:
        print(f"  <{el.name} class='{el.get('class','')}'>  {t[:150]}")

print()
print("=" * 60)
print("TEXTE CONTENANT 'ann' + année :")
print("=" * 60)
for el in soup.find_all(True):
    t = el.get_text(separator=" ", strip=True)
    if re.search(r"ann[ée]e\s*:\s*\d{4}", t, re.I) and len(t) < 200:
        print(f"  <{el.name} class='{el.get('class','')}'>  {t[:150]}")
