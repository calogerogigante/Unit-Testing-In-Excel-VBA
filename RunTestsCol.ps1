# ============================================================
# RunTestsCol.ps1 - lance les tests unitaires VBA d'un classeur Excel (AVEC COULEURS)
#
# Usage (depuis le dossier du classeur) :
#   powershell -ExecutionPolicy Bypass -File .\RunTestsCol.ps1
#       -> teste le premier .xlsm trouve dans le dossier du script
#   powershell -ExecutionPolicy Bypass -File .\RunTestsCol.ps1 -Classeur "C:\chemin\Tests.xlsm"
#       -> teste un classeur precis
#   Clic droit sur le script > "Executer avec PowerShell"
#       -> la fenetre reste ouverte jusqu'a ce qu'on appuie sur une touche
#   Ajouter -SansPause pour ne PAS attendre de touche (usage en terminal ou automatise)
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
    [string]$Classeur,   # optionnel : sans argument, teste le 1er .xlsm du dossier du script
    [switch]$SansPause   # ne pas attendre une touche a la fin (terminal, automatisation)
)

# ============================================================
# CONFIGURATION DE L'AFFICHAGE (palette 1, niveau 1)
# Les couleurs RGB (0 a 255) sont utilisees si la console gere les codes ANSI ;
# sinon le script retombe sur les couleurs de console classiques.
# ============================================================
$Titre = "Tests unitaires sur code VBA (user prenom-nom)"
$TitreRgb = @(255, 235, 90)             # jaune clair
$TitreRepli = 'Yellow'

# Etiquettes [OK] [ECHEC] [ERREUR] : couleur du texte et FOND PLEIN (couleurs de console)
$EtiquetteFg = @{ OK = 'Black'; ECHEC = 'White'; ERREUR = 'Black' }
$EtiquetteBg = @{ OK = 'Green'; ECHEC = 'Red';   ERREUR = 'Yellow' }

# Texte (nom du test, message, bilan) : RGB, et couleur de console de repli
$TexteRgb     = @{ OK = @(110, 255, 110); ECHEC = @(255, 105, 105); ERREUR = @(255, 235, 90) }
$TexteConsole = @{ OK = 'Green';          ECHEC = 'Red';            ERREUR = 'Yellow' }

# --- Active l'interpretation des codes ANSI dans la console classique (Windows 10+) ---
$ansiOk = $false
try
{
    if (-not ([System.Management.Automation.PSTypeName]'Win32.ConsoleVT').Type)
    {
        Add-Type -Namespace Win32 -Name ConsoleVT -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int nStdHandle);
[DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
[DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
'@
    }

    $handle = [Win32.ConsoleVT]::GetStdHandle(-11)
    $mode = [uint32]0
    if ([Win32.ConsoleVT]::GetConsoleMode($handle, [ref]$mode))
    {
        # 4 = ENABLE_VIRTUAL_TERMINAL_PROCESSING
        $ansiOk = [Win32.ConsoleVT]::SetConsoleMode($handle, ($mode -bor 4))
    }
}
catch { }

$reset = "$([char]27)[0m"

function Ansi($c)
{
    return "$([char]27)[38;2;$($c[0]);$($c[1]);$($c[2])m"
}

# Ligne de resultat : etiquette en fond plein + nom du test + message eventuel
function Ecrire-Resultat([string]$etat, [string]$nom, [string]$message)
{
    $tag = "[$etat]"
    Write-Host (" " + $tag.PadRight(9)) -NoNewline -ForegroundColor ($EtiquetteFg[$etat]) -BackgroundColor ($EtiquetteBg[$etat])

    $texte = " " + $nom
    if ($message) { $texte += " : " + $message }

    if ($ansiOk)
    {
        Write-Host ((Ansi ($TexteRgb[$etat])) + $texte + $reset)
    }
    else
    {
        Write-Host $texte -ForegroundColor ($TexteConsole[$etat])
    }
}

# Morceau de texte colore selon l'etat (pour la ligne de bilan)
function Ecrire-Segment([string]$texte, [string]$etat, [bool]$fin)
{
    if ($ansiOk)
    {
        $s = (Ansi ($TexteRgb[$etat])) + $texte + $reset
        if ($fin) { Write-Host $s } else { Write-Host $s -NoNewline }
    }
    else
    {
        if ($fin) { Write-Host $texte -ForegroundColor ($TexteConsole[$etat]) }
        else      { Write-Host $texte -NoNewline -ForegroundColor ($TexteConsole[$etat]) }
    }
}

# Titre en jaune clair, suivi d'une ligne de tirets
function Ecrire-Titre
{
    $ligne = "-" * 64

    if ($ansiOk)
    {
        Write-Host ((Ansi $TitreRgb) + $Titre + $reset)
        Write-Host ((Ansi $TitreRgb) + $ligne + $reset)
    }
    else
    {
        Write-Host $Titre -ForegroundColor $TitreRepli
        Write-Host $ligne -ForegroundColor $TitreRepli
    }
}

# Attend une touche avant de fermer la fenetre (sauf avec -SansPause)
function Attendre-Touche
{
    if ($SansPause) { return }

    Write-Host ""
    Write-Host "Appuyez sur une touche pour fermer cette fenetre..." -ForegroundColor DarkGray

    try
    {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    catch
    {
        # Environnement sans ReadKey (ex. PowerShell ISE) : on attend Entree
        $null = Read-Host "Appuyez sur Entree"
    }
}

# Filet de securite : toute erreur imprevue s'affiche et la fenetre reste ouverte
trap
{
    Write-Host ""
    Write-Host "ERREUR : $($_.Exception.Message)" -ForegroundColor Red
    Attendre-Touche
    exit 1
}

Ecrire-Titre

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
                    Ecrire-Resultat "OK" $t ""
                }
                elseif ($parts[0] -eq "KO")
                {
                    $echecs++
                    Ecrire-Resultat "ECHEC" $t $parts[1]
                }
                else
                {
                    $erreurs++
                    Ecrire-Resultat "ERREUR" $t $parts[1]
                }
            }
            catch
            {
                $erreurs++
                Ecrire-Resultat "ERREUR" $t $_.Exception.Message
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
Write-Host "$total tests : " -NoNewline
Ecrire-Segment "$ok OK" "OK" $false
Write-Host ", " -NoNewline
Ecrire-Segment "$echecs ECHEC" "ECHEC" $false
Write-Host ", " -NoNewline
Ecrire-Segment "$erreurs ERREUR" "ERREUR" $true

Attendre-Touche

if ($echecs + $erreurs -eq 0) { exit 0 } else { exit 1 }
