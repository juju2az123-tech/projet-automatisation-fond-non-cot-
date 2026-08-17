#!/usr/bin/env python3
"""
Remplace, dans les pages HTML de l'outil, les <script src="..."></script> pointant vers des
fichiers locaux par leur contenu inliné directement dans le HTML.

Pourquoi : quand un fichier est distribué en .zip, un simple double-clic sur le .html DEPUIS
l'aperçu de l'explorateur Windows (sans "Extraire tout" au préalable) n'extrait que ce seul
fichier dans un dossier temporaire — les sous-dossiers data/, js/, vendor/ ne sont pas présents
à côté, et les <script src="..."> échouent silencieusement (l'app se charge mais reste inerte).
Un fichier HTML 100% autonome (aucune dépendance externe) élimine ce mode d'échec.

Les fichiers *.template.html sont la source (éditez ceux-là, avec leurs <script src="...">
normaux). Ce script les lit et écrit la version finale, autonome, sans le suffixe ".template" —
c'est celle-là qu'on distribue/utilise. Ne jamais éditer directement index.html ou
ajout-calendrier-sortie.html : ils sont régénérés à chaque exécution et vos changements seraient
perdus.

Usage :
    python3 scripts/inline_assets.py
Régénère index.html à partir de index.template.html, et ajout-calendrier-sortie.html à partir de
ajout-calendrier-sortie.template.html. À relancer après toute modification d'un template ou de
data/funds_data.js / js/exit_calendar_builder.js / vendor/exceljs.min.js.
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

TEMPLATES = ["index.template.html", "ajout-calendrier-sortie.template.html"]

SCRIPT_SRC_RE = re.compile(r'<script src="([^"]+)"></script>')


def inline_one(template_path: Path, out_path: Path):
    html = template_path.read_text(encoding="utf-8")

    def repl(m):
        rel = m.group(1)
        asset_path = ROOT / rel
        content = asset_path.read_text(encoding="utf-8")
        # Guard against a stray "</script>" inside the asset breaking the wrapper tag.
        content = content.replace("</script>", "<\\/script>")
        return f'<script>\n{content}\n</script>'

    new_html, n = SCRIPT_SRC_RE.subn(repl, html)
    out_path.write_text(new_html, encoding="utf-8")
    print(f"{out_path.name}: {n} <script src> inliné(s) (depuis {template_path.name})")


def main():
    for name in TEMPLATES:
        p = ROOT / name
        if not p.exists():
            print(f"(ignoré, introuvable) {name}")
            continue
        out_name = name.replace(".template.html", ".html")
        inline_one(p, ROOT / out_name)


if __name__ == "__main__":
    main()
