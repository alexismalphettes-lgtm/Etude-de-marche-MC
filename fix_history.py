"""
Script de migration unique — corrige les prix corrompus dans boats_history.json.

Contexte : l'ancienne version de parse_price traitait mal la virgule décimale
française. Ex: "495 000,00 €" donnait 49 500 000 au lieu de 495 000 (x100).

Règle de correction :
  - Si un prix stocké est > 5 000 000 ET que price/100 donne un prix cohérent
    (entre 1 000 et 5 000 000 €), on divise par 100.
  - Sinon on laisse intact.

Aucune donnée n'est supprimée, les baisses de prix légitimes sont conservées.
"""

import json
from pathlib import Path

DATA_FILE = Path("data/boats_history.json")
THRESHOLD = 5_000_000   # prix > ce seuil = suspect
MIN_PRICE  = 1_000      # prix corrigé minimum plausible
MAX_PRICE  = 5_000_000  # prix corrigé maximum plausible (AN66 vend Jeanneau/Prestige)


def fix_price(price):
    if price is None:
        return price
    if price > THRESHOLD:
        corrected = price / 100
        if MIN_PRICE <= corrected <= MAX_PRICE:
            return round(corrected, 2)
    return price


def main():
    if not DATA_FILE.exists():
        print("Aucun fichier historique trouvé — rien à corriger.")
        return

    with open(DATA_FILE, "r", encoding="utf-8") as f:
        history = json.load(f)

    total_fixed = 0

    for category in ("neufs", "occasions"):
        for url, boat in history.get(category, {}).items():
            for entry in boat.get("historique_prix", []):
                old = entry["prix"]
                new = fix_price(old)
                if new != old:
                    print(f"  [{category}] {boat.get('modele','?'):<35} "
                          f"{old:>15,.0f} € → {new:>12,.0f} €")
                    entry["prix"] = new
                    total_fixed += 1

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    print(f"\n✓ {total_fixed} prix corrigés — historique sauvegardé.")


if __name__ == "__main__":
    main()
