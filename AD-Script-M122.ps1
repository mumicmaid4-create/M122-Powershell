#Requires -Version 3.0
 
<#
Variablenübersicht:
 
  $labels         - Array mit den Feldnamen für die Eingabefelder im Formular
  $textBoxes      - Hashtable zum Speichern der TextBox-Objekte
  $vorname        - Eingegebener Vorname des neuen Benutzers
  $nachname       - Eingegebener Nachname des neuen Benutzers
  $sAMAccountName - Eingegebener Login-Name (sAMAccountName)
  $DisplayName    - Eingegebner Anzeigename
  $department     - Eingegebene Abteilung
  $passwordPlain  - Eingegebenes Klartext-Passwort
  $description    - Eingegebene Beschreibung für das Benutzerkonto
#>
 
 
# Lädt die benötigten Komponenten
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
 
# Formular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "Neuer AD-Benutzer"
$form.Size = New-Object System.Drawing.Size(500,360)
$form.StartPosition = "CenterScreen"
 
# Beschriftungen und Eingabefelder
$labels = @("Vorname","Nachname","Displayname", "Anmeldename","Kennwort","Jobtitel", "Abteilung")
$y = 20
$textBoxes = @{}
foreach ($labelText in $labels) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($labelText):"
    $lbl.Location = New-Object System.Drawing.Point(10,$y)
    $lbl.Size = New-Object System.Drawing.Size(100,20)
    $form.Controls.Add($lbl)
 
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(150,$y)
    $txt.Size = New-Object System.Drawing.Size(250,20)
    if ($labelText -eq "Kennwort") { $txt.UseSystemPasswordChar = $true }
    $form.Controls.Add($txt)
    $textBoxes[$labelText] = $txt
 
    $y += 30
}
 
# OK-Button ohne DialogResult, validierung im Click-Event
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(200, $y)
$okButton.AutoSize = $true
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)
 
# Abbrechen-Button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Abbrechen"
$cancelButton.Location = New-Object System.Drawing.Point(280, $y)
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
 
# Validierung und Anlegen
$okButton.Add_Click({
    # Werte holen
    $vorname        = $textBoxes["Vorname"].Text.Trim()
    $nachname       = $textBoxes["Nachname"].Text.Trim()
    $sAMAccountName = $textBoxes["Anmeldename"].Text.Trim()
    $DisplayName    = $textBoxes["Displayname"].Text.Trim()
    $passwordPlain  = $textBoxes["Kennwort"].Text
    $description    = $textBoxes["Jobtitel"].Text.Trim()
    $department     = $textBoxes["Abteilung"].Text.Trim()
 
    # Funktion zur Passwort-Komplexitätsprüfung
    function Test-PasswordComplexity {
    param([string]$pwd)
    return $pwd.Length -ge 14 -and
           $pwd -match '[A-Z]' -and
           $pwd -match '[a-z]' -and
           $pwd -match '\d'   -and
           $pwd -match '[^A-Za-z0-9]'
    }
 
     # Passwort-Komplexität prüfen
    if (-not (Test-PasswordComplexity -pwd $passwordPlain)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Das Passwort erfüllt nicht die Komplexitätsanforderungen.`Es muss mindestens 14 Zeichen lang sein und Groß- und Kleinbuchstaben, Zahlen sowie Sonderzeichen enthalten.",
            "Ungültiges Passwort",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
 
    # Feldvalidierung
    if ([string]::IsNullOrWhiteSpace($vorname) -or
        [string]::IsNullOrWhiteSpace($nachname) -or
        [string]::IsNullOrWhiteSpace($sAMAccountName) -or
        [string]::IsNullOrWhiteSpace($DisplayName) -or
        [string]::IsNullOrWhiteSpace($passwordPlain) -or
        [string]::IsNullOrWhiteSpace($department) -or
        [string]::IsNullOrWhiteSpace($description)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte füllen Sie alle Felder aus.",
            "Fehlende Eingabe",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
 
   
 
    # Passwort in SecureString umwandeln
    $password = ConvertTo-SecureString $passwordPlain -AsPlainText -Force
 
    # Benutzer anlegen mit Fehlerbehandlung
    try {
        New-ADUser `
            -Name "$vorname $nachname" `
            -DisplayName $DisplayName `
            -GivenName $vorname `
            -Surname $nachname `
            -SamAccountName $sAMAccountName `
            -Description $description `
            -AccountPassword $password `
            -UserPrincipalName  "$sAMAccountName@Mumic.ch" `
            -EmailAddress "$vorname.$Nachname@Mumic.ch" `
            -Office "Zürich" `
            -Street "Ausstelungs-strasse 15" `
            -PostalCode "8001" `
            -City "Zürich" `
            -Company "Mumic" `
            -State "Zürich" `
            -Country "CH" `
            -Title $description `
            -Department $department `
            -Enabled $true `
            -ChangePasswordAtLogon $true `
            -Path "OU=Users,OU=Zürich,DC=Mumic,DC=ch"`
            -ErrorAction Stop
 
         # Erfolgsmeldung
        [System.Windows.Forms.MessageBox]::Show(
            "Benutzer '$sAMAccountName' erfolgreich erstellt.",
            "Erfolg",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)

 
        # Summary-Inhalt erstellen
        $datum = Get-Date -Format "dd-MM-yyyy"
        $summaryFile = Join-Path -Path "C:\IT\Neue_Mitarbeiter" -ChildPath "${datum}_${sAMAccountName}_Summary.txt"
        $lines = @(
            "Willkommen im Unternehmen, $vorname $nachname!",
            "Ihr neues Benutzerkonto wurde erstellt. Wir bitten Sie, die Details zu 
            überprüfen und Fehlehrhaft Einträge dem HR zu melden:",
            "  - Vorname: $Vorname",
            "  - Nachname: $nachname",
            "  - Benutzername: $sAMAccountName",
            "  - Beschreibung: $description",
            "  - Ihr Passwort: $passwordPlain" ,
            "Bitte ändern Sie bei der ersten Anmeldung Ihr Kennwort."          
        )
        $lines | Out-File -FilePath $summaryFile
 
        # Formular schließen
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler: $($_.Exception.Message)",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
 
# Formular anzeigen
$form.ShowDialog() | Out-Null
