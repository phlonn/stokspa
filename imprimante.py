"""Impression directe d'une liste de prélèvement (classeur Excel) sur l'imprimante par défaut de la machine
où tourne l'application (Windows + Excel installés). Aucune fenêtre ne s'ouvre : Excel imprime en arrière-plan.

Avant l'impression, la feuille est mise en page : en haut les informations de la livraison
(n° de livraison, n° de voyage, client, transporteur…), puis le tableau des produits en grands caractères."""

from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import tempfile

DISPONIBLE = sys.platform == "win32"
UNE_PAGE = 20  # nombre de lignes d'emplacement qui tiennent sur une seule page

_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$xl = $null
try {
  try { $xl = New-Object -ComObject Excel.Application } catch { Write-Output 'PASEXCEL'; exit 0 }
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  $infos = $env:IMP_INFOS | ConvertFrom-Json
  $wb = $xl.Workbooks.Open($env:IMP_FICHIER)
  $ws = $wb.Worksheets.Item(1)
  if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
  $nCols = $ws.UsedRange.Columns.Count
  $nRows = $ws.UsedRange.Rows.Count
  $lastCol = [char](64 + $nCols)

  # --- Tableau des produits : grands caractères, quadrillé ---
  $tab = $ws.Range("A1:$lastCol$nRows")
  $tab.Font.Size = [int]$env:IMP_TAILLE
  $tab.VerticalAlignment = -4108
  foreach ($b in 7, 8, 9, 10, 11, 12) { $tab.Borders.Item($b).LineStyle = 1; $tab.Borders.Item($b).Weight = 2; $tab.Borders.Item($b).Color = 0 }
  $ws.Range("A1:$($lastCol)1").Font.Bold = $true
  $ws.Range("A1:$($lastCol)1").Interior.Color = 0xE0E0E0
  $tab.Rows.RowHeight = [double]$env:IMP_HAUTEUR
  # colonnes Emplacement et Palettes centrées
  for ($c = 1; $c -le $nCols; $c++) {
    if (@('Emplacement', 'Palettes') -contains [string]$ws.Cells.Item(1, $c).Value2) {
      $ws.Range($ws.Cells.Item(1, $c), $ws.Cells.Item($nRows, $c)).HorizontalAlignment = -4108
    }
  }
  $ws.Columns.AutoFit() | Out-Null
  for ($c = 1; $c -le $nCols; $c++) { $col = $ws.Columns.Item($c); $col.ColumnWidth = $col.ColumnWidth + 3 }

  # --- Informations de la livraison au-dessus du tableau ---
  $n = @($infos.lignes).Count
  $ws.Rows("1:$($n + 3)").Insert() | Out-Null
  $ws.Range("A1:$($lastCol)1").Font.Size = 24
  $ws.Range("A1:$($lastCol)1").Interior.ColorIndex = -4142
  $ws.Range("A1:$($lastCol)$($n + 3)").Borders.LineStyle = -4142
  $ws.Cells.Item(1, 1).Value2 = $infos.titre
  $ws.Cells.Item(1, 1).Font.Bold = $true
  $ws.Rows(1).RowHeight = 36
  $r = 3
  foreach ($l in @($infos.lignes)) {
    $grand = (@($l).Count -gt 2) -and $l[2]   # information mise en évidence (ex. n° de voyage)
    $ws.Range("A$($r):$($lastCol)$($r)").Font.Size = $(if ($grand) { 44 } else { 16 })
    $ws.Range("A$($r):$($lastCol)$($r)").Font.Bold = $false
    $ws.Range("A$($r):$($lastCol)$($r)").Interior.ColorIndex = -4142
    $ws.Cells.Item($r, 1).Value2 = $l[0] + ' :'
    $ws.Cells.Item($r, 2).Value2 = "'" + $l[1]
    $ws.Cells.Item($r, 2).Font.Bold = $true
    if ($grand) { $ws.Cells.Item($r, 1).Font.Bold = $true; $ws.Cells.Item($r, 1).Font.Size = 18; $ws.Cells.Item($r, 1).VerticalAlignment = -4108 }
    $ws.Cells.Item($r, 2).HorizontalAlignment = -4131
    $ws.Rows($r).RowHeight = $(if ($grand) { 62 } else { 26 })
    $r++
  }
  $ws.Rows($r).RowHeight = 14
  $ws.Range("A$($r):$($lastCol)$($r)").Interior.ColorIndex = -4142
  # libellés sur les colonnes A+B, valeurs sur C et suivantes : la largeur des colonnes du tableau n'en dépend pas
  $maxLab = 0
  for ($k = 3; $k -lt $r; $k++) {
    $lab = [string]$ws.Cells.Item($k, 1).Value2; $val = $ws.Cells.Item($k, 2).Value2
    $sz = [double]$ws.Cells.Item($k, 1).Font.Size
    $maxLab = [Math]::Max([double]$maxLab, [double]$lab.Length * $sz / 11.0 * 1.15)
    $ws.Cells.Item($k, 2).ClearContents() | Out-Null
    $ws.Cells.Item($k, 3).Value2 = "'" + $val
    $ws.Cells.Item($k, 3).Font.Bold = $true   # la taille est déjà celle de la ligne
    $ws.Cells.Item($k, 3).HorizontalAlignment = -4131   # textes qui débordent sur les cases vides voisines
  }
  $ab = $ws.Columns.Item(1).ColumnWidth + $ws.Columns.Item(2).ColumnWidth
  if ($ab -lt $maxLab) { $ws.Columns.Item(2).ColumnWidth = $ws.Columns.Item(2).ColumnWidth + ($maxLab - $ab) }

  # --- Page : portrait, largeur d'une page, en-tête de tableau répété ---
  $ps = $ws.PageSetup
  $ps.Orientation = 1
  $ps.Zoom = $false; $ps.FitToPagesWide = 1
  if ($env:IMP_UNEPAGE) { $ps.FitToPagesTall = 1 } else { $ps.FitToPagesTall = $false }
  $ps.PrintTitleRows = '$' + ($n + 4) + ':$' + ($n + 4)
  $ps.CenterHorizontally = $true
  $ps.LeftMargin = $xl.CentimetersToPoints(1); $ps.RightMargin = $xl.CentimetersToPoints(1)
  $ps.TopMargin = $xl.CentimetersToPoints(1.5); $ps.BottomMargin = $xl.CentimetersToPoints(1.5)
  $ps.CenterFooter = '&12Page &P / &N'

  if ($env:IMP_PDF) { $wb.ExportAsFixedFormat(0, $env:IMP_PDF); Write-Output 'OK PDF' }
  else { $ws.PrintOut(); Write-Output ('OK ' + $xl.ActivePrinter) }
  $wb.Close($false)
} catch {
  Write-Output ('ERREUR ' + $_.Exception.Message + ' (ligne ' + $_.InvocationInfo.ScriptLineNumber + ')')
} finally {
  if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}
"""


class ImpressionError(Exception):
    pass


def imprimer_xlsx(donnees_b64: str, titre: str = "", lignes: list | None = None, taille: int = 20, pdf: str = "") -> str:
    """Met en page et imprime la première feuille du classeur ; renvoie le nom de l'imprimante utilisée.
    lignes : [[libellé, valeur], …] affichées au-dessus du tableau. pdf : chemin d'un PDF à produire au lieu d'imprimer (essai)."""
    if not DISPONIBLE:
        raise ImpressionError("L'impression directe n'est disponible que sur l'application installée sur un PC Windows.")
    fd, fichier = tempfile.mkstemp(suffix=".xlsx", prefix="impression_")
    with os.fdopen(fd, "wb") as f:
        f.write(base64.b64decode(donnees_b64))
    # Jusqu'à UNE_PAGE lignes d'emplacement : tout tient sur une seule page A4 (caractères et hauteur de ligne
    # réduits si besoin) ; au-delà, la liste continue sur plusieurs pages en grands caractères.
    try:
        import io, re, zipfile
        feuille = zipfile.ZipFile(io.BytesIO(base64.b64decode(donnees_b64))).read("xl/worksheets/sheet1.xml").decode("utf-8")
        n_lignes = max(0, len(re.findall(r"<row ", feuille)) - 1)
    except Exception:
        n_lignes = 0
    une_page = 0 < n_lignes <= UNE_PAGE
    hauteur = min(2.0 * taille, 480 / (n_lignes + 1)) if une_page else 2.0 * taille
    taille = min(taille, max(11, int(hauteur / 1.7)))
    infos = {"titre": titre[:120],
             "lignes": [[str(l[0])[:40], str(l[1])[:120], 1 if len(l) > 2 and l[2] else 0] for l in (lignes or []) if len(l) > 1 and str(l[1]).strip()]}
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", _SCRIPT],
            capture_output=True, text=True, timeout=180,
            env={**os.environ, "IMP_FICHIER": fichier, "IMP_INFOS": json.dumps(infos, ensure_ascii=False),
                 "IMP_TAILLE": str(int(taille)), "IMP_HAUTEUR": f"{hauteur:.1f}", "IMP_UNEPAGE": "1" if une_page else "",
                 "IMP_PDF": pdf},
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        sortie = (r.stdout or "").strip().splitlines()
        res = sortie[-1] if sortie else (r.stderr or "").strip()
        if res == "PASEXCEL":
            raise ImpressionError("Excel n'est pas installé sur ce PC : utilisez « Liste en Excel » puis imprimez le fichier.")
        if not res.startswith("OK"):
            raise ImpressionError("Impression impossible : " + (res.removeprefix("ERREUR ") or "erreur inconnue"))
        return res[3:].split(" sur ")[0].split(" on ")[0].strip()
    except subprocess.TimeoutExpired:
        raise ImpressionError("L'impression a pris trop de temps et a été arrêtée.")
    finally:
        try:
            os.remove(fichier)
        except OSError:
            pass
