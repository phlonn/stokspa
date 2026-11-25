"""Numérisation avec le scanner branché sur la machine où tourne l'application (Windows uniquement).

Ouvre la fenêtre de numérisation de Windows (WIA), qui fonctionne avec tous les scanners installés,
et renvoie la page numérisée en JPEG (base64)."""

from __future__ import annotations

import base64
import os
import subprocess
import sys
import tempfile

DISPONIBLE = sys.platform == "win32"

# Format JPEG de WIA
_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"

_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
try {
  $dlg = New-Object -ComObject WIA.CommonDialog
  # 1 = scanner ; fenêtre de numérisation de Windows ; pas d'erreur si on annule
  $img = $dlg.ShowAcquireImage(1, 0, 131072, '%s', $false, $true, $false)
  if ($img -eq $null) { Write-Output 'ANNULE'; exit 0 }
  if ($img.FormatID -ne '%s') {
    $ip = New-Object -ComObject WIA.ImageProcess
    $ip.Filters.Add($ip.FilterInfos.Item('Convert').FilterID)
    $ip.Filters.Item(1).Properties.Item('FormatID').Value = '%s'
    $ip.Filters.Item(1).Properties.Item('Quality').Value = 90
    $img = $ip.Apply($img)
  }
  if (Test-Path -LiteralPath $env:SCAN_SORTIE) { Remove-Item -LiteralPath $env:SCAN_SORTIE }
  $img.SaveFile($env:SCAN_SORTIE)
  Write-Output 'OK'
} catch {
  $m = $_.Exception.Message
  if ($m -match '0x80210015' -or $m -match 'no WIA device') { Write-Output 'AUCUN' } else { Write-Output ('ERREUR ' + $m) }
}
""" % (_JPEG, _JPEG, _JPEG)


class ScanError(Exception):
    pass


def numeriser() -> str | None:
    """Renvoie l'image numérisée (JPEG en base64), ou None si l'utilisateur a annulé."""
    if not DISPONIBLE:
        raise ScanError("Le scanner n'est disponible que sur l'application installée sur un PC Windows.")
    fd, sortie = tempfile.mkstemp(suffix=".jpg", prefix="scan_")
    os.close(fd)
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", _SCRIPT],
            capture_output=True, text=True, timeout=900, env={**os.environ, "SCAN_SORTIE": sortie},
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        res = (r.stdout or "").strip().splitlines()
        res = res[-1] if res else (r.stderr or "").strip()
        if res == "ANNULE":
            return None
        if res == "AUCUN":
            raise ScanError("Aucun scanner n'a été trouvé sur ce PC. Vérifiez qu'il est allumé et branché.")
        if res != "OK":
            raise ScanError("Numérisation impossible : " + (res.removeprefix("ERREUR ") or "erreur inconnue"))
        with open(sortie, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except subprocess.TimeoutExpired:
        raise ScanError("La numérisation a pris trop de temps et a été arrêtée.")
    finally:
        try:
            os.remove(sortie)
        except OSError:
            pass
