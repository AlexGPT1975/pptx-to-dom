"""Referenzbilder einer PPTX, gerendert von Microsoft PowerPoint (macOS).

    python bench/referenz.py datei.pptx zielordner/

PowerPoint ist sandboxed: Die Datei wird in seinen Container kopiert, per AppleScript unter
ihrem NAMEN geöffnet (nie „active presentation" - andere offene Präsentationen bleiben
unberührt), als PDF gespeichert (der PNG-Export per AppleScript schreibt auf macOS nichts)
und geschlossen. pdftoppm rastert das PDF mit 144 dpi: 13,333 x 7,5 Zoll = 1920 x 1080.
Ergebnis: zielordner/folie_01.png ...
"""
from __future__ import annotations

import pathlib
import shutil
import subprocess
import sys
import time

CONTAINER = pathlib.Path.home() / "Library/Containers/com.microsoft.Powerpoint/Data/oc_import"


def _osa(skript: str) -> str:
    r = subprocess.run(["osascript", "-"], input=skript, capture_output=True, text=True, timeout=300)
    if r.returncode:
        raise RuntimeError(f"AppleScript: {r.stderr.strip()}")
    return r.stdout.strip()


def pdf_von_powerpoint(pptx: pathlib.Path) -> pathlib.Path:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    name = f"oc_rundlauf_{pptx.stem}.pptx"
    kopie, pdf = CONTAINER / name, CONTAINER / f"oc_rundlauf_{pptx.stem}.pdf"
    shutil.copyfile(pptx, kopie)
    pdf.unlink(missing_ok=True)
    _osa(f'''
tell application "Microsoft PowerPoint"
  if (name of every presentation) contains "{name}" then close presentation "{name}" saving no
  open POSIX file "{kopie}"
  repeat 40 times
    if (name of every presentation) contains "{name}" then exit repeat
    delay 0.5
  end repeat
  set p to presentation "{name}"
  save p in ((POSIX file "{pdf}") as text) as save as PDF
  close p saving no
end tell''')
    for _ in range(60):
        if pdf.exists() and pdf.stat().st_size > 0:
            break
        time.sleep(0.5)
    if not pdf.exists():
        raise RuntimeError("PowerPoint hat kein PDF geschrieben")
    return pdf


def rastern(pdf: pathlib.Path, ziel: pathlib.Path, dpi: int = 144) -> list[pathlib.Path]:
    ziel.mkdir(parents=True, exist_ok=True)
    for alt in ziel.glob("folie_*.png"):
        alt.unlink()
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(ziel / "seite")], check=True)
    out = []
    for i, f in enumerate(sorted(ziel.glob("seite*.png")), 1):
        neu = ziel / f"folie_{i:02d}.png"
        f.replace(neu)
        out.append(neu)
    return out


def referenz(pptx: pathlib.Path, ziel: pathlib.Path) -> list[pathlib.Path]:
    pdf = pdf_von_powerpoint(pptx)
    shutil.copyfile(pdf, ziel.parent / f"{pptx.stem}.powerpoint.pdf") if ziel.parent.exists() else None
    return rastern(pdf, ziel)


if __name__ == "__main__":
    for f in referenz(pathlib.Path(sys.argv[1]).resolve(), pathlib.Path(sys.argv[2]).resolve()):
        print(f)
