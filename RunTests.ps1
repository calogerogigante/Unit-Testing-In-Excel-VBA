# ============================================================
# RunTests.ps1 - lance les tests unitaires VBA d'un classeur Excel
#
# Usage :
#   - clic droit sur le fichier > Executer avec PowerShell
#       -> teste le premier .xlsm trouve dans le dossier du script
#   - ou en ligne de commande :
#       powershell -ExecutionPolicy Bypass -File .\RunTests.ps1 -Classeur "C:\chemin\Tests.xlsm"
#   - ajouter -SansPause pour ne PAS attendre une touche a la fin
#
# Prerequis :
#   - Excel installe (Windows)
#   - Excel > Options > Centre de gestion de la confidentialite > Parametres des macros
#     > cocher "Acces approuve au modele d'objet du projet VBA"
#   - Le classeur doit etre ferme dans Excel pendant l'execution
#
# Trois statuts par test : [OK], [ECHEC] (assertion fausse), [ERREUR] (le code a plante)
# ============================================================
param(
    [string]$Classeur,    # optionnel : sans argument, teste le 1er .xlsm du dossier du script
    [switch]$SansPause    # ne pas attendre une touche a la fin
)

# Attend une touche pour que la fenetre ne se ferme pas toute seule
function Attendre-Touche
{
    if ($SansPause) { return }
    Write-Host ""
    Write-Host "Appuyez sur une touche pour fermer cette fenetre..." -ForegroundColor DarkGray
    try   { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") }
    catch { $null = Read-Host "Appuyez sur Entree" }
}

# Filet de securite : toute erreur imprevue s'affiche et la fenetre reste ouverte
trap
{
    Write-Host ""
    Write-Host "ERREUR : $($_.Exception.Message)" -ForegroundColor Red
    Attendre-Touche
    exit 1
}

if ($Classeur)
{
    $chemin = (Resolve-Path $Classeur).Path
}
else
{
    # Premier .xlsm (ordre alphabetique) du dossier du script, en ignorant les
    # fichiers temporaires "~$..." qu'Excel cree quand un classeur est ouvert
    $fichier = Get-ChildItem -Path $PSScriptRoot -Filter "*.xlsm" |
               Where-Object { $_.Name -notlike '~$*' } |
               Sort-Object Name |
               Select-Object -First 1

    if (-not $fichier)
    {
        Write-Host "Aucun fichier .xlsm trouve dans $PSScriptRoot" -ForegroundColor Red
        Attendre-Touche
        exit 2
    }
    $chemin = $fichier.FullName
}

Write-Host "Classeur teste : $chemin" -ForegroundColor DarkGray

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow : autorise les macros a l'ouverture

$total = 0
$ok = 0
$echecs = 0
$erreurs = 0
$wb = $null

# Appelle une macro "Module.Procedure" du classeur et retourne sa valeur
function Appeler([string]$module, [string]$proc)
{
    return $excel.Run("'$($wb.Name)'!$module.$proc")
}

# Indique si le code contient un "Public Sub <nom>"
function ADeclare([string]$code, [string]$nom)
{
    return [regex]::IsMatch($code, "(?im)^\s*Public\s+Sub\s+$nom\b")
}

try
{
    $wb = $excel.Workbooks.Open($chemin)

    foreach ($comp in $wb.VBProject.VBComponents)
    {
        # Type 1 = module standard ; on ne garde que ceux dont le nom commence par "Test"
        if ($comp.Type -ne 1 -or $comp.Name -notlike "Test*") { continue }

        $cm = $comp.CodeModule
        if ($cm.CountOfLines -eq 0) { continue }
        $code = $cm.Lines(1, $cm.CountOfLines)

        $tests = @([regex]::Matches($code, '(?im)^\s*Public\s+Function\s+(Test\w*)') |
                   ForEach-Object { $_.Groups[1].Value })

        Write-Host ""
        Write-Host "== $($comp.Name) ($($tests.Count) tests)" -ForegroundColor Cyan

        $aBeforeAll  = ADeclare $code "BeforeAll"
        $aBeforeEach = ADeclare $code "BeforeEach"
        $aAfterEach  = ADeclare $code "AfterEach"
        $aAfterAll   = ADeclare $code "AfterAll"

        if ($aBeforeAll) { [void](Appeler $comp.Name "BeforeAll") }

        foreach ($t in $tests)
        {
            $total++
            try
            {
                if ($aBeforeEach) { [void](Appeler $comp.Name "BeforeEach") }

                # Le test est execute cote Excel par RunnerHelper, qui renvoie "STATUT|message"
                $nomTest = "'$($wb.Name)'!$($comp.Name).$t"
                $res = [string]$excel.Run("'$($wb.Name)'!RunnerHelper.ExecuterTest", $nomTest)
                $parts = $res.Split("|", 2)

                if ($aAfterEach) { [void](Appeler $comp.Name "AfterEach") }

                if ($parts[0] -eq "OK")
                {
                    $ok++
                    Write-Host "  [OK]     $t" -ForegroundColor Green
                }
                elseif ($parts[0] -eq "KO")
                {
                    $echecs++
                    Write-Host "  [ECHEC]  $t : $($parts[1])" -ForegroundColor Red
                }
                else
                {
                    $erreurs++
                    Write-Host "  [ERREUR] $t : $($parts[1])" -ForegroundColor Yellow
                }
            }
            catch
            {
                $erreurs++
                Write-Host "  [ERREUR] $t : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if ($aAfterAll) { [void](Appeler $comp.Name "AfterAll") }
    }
}
finally
{
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

Write-Host ""
if ($echecs + $erreurs -eq 0)
{
    Write-Host "$total tests : $ok OK, $echecs ECHEC, $erreurs ERREUR" -ForegroundColor Green
}
else
{
    Write-Host "$total tests : $ok OK, $echecs ECHEC, $erreurs ERREUR" -ForegroundColor Red
}

Attendre-Touche

if ($echecs + $erreurs -eq 0) { exit 0 } else { exit 1 }
