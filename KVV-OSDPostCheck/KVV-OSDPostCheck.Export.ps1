Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Create-Form {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Kriminalvården Deployment Compliance"
    $form.Size = New-Object System.Drawing.Size(500, 400)  # Updated size
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog  # Prevent resizing
    $form.MaximizeBox = $false  # Disable maximize button
    $form.MinimizeBox = $false  # Disable minimize button
    # Base64 encoded icon string
    $iconBase64 = @"
    AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAACMuAAAjLgAAAAAAAAAAAAD////////////////////////////////////////////////v7Ov/0sW7/7WWgP+eb03/kFQo/4hEEv+FPQf/hT0H/4hEEv+QVCj/nnBN/7WXgf/Txbv/7+3r//////////////////////////////////////////////////////////////////////////////////3+/v/h2dT/tpeB/5hgN/+NRA//jDwB/407AP+PPAD/jz0A/489AP+PPQD/jz0A/488AP+NOwD/jDwB/41FD/+YYTj/tpiC/+Ha1f/9/v/////////////////////////////////////////////////////////////s6ef/spN9/49NHf+LOgD/jjsA/489AP+PPQD/jzwA/488AP+PPQD/jzwA/488AP+PPQD/jz0A/488AP+PPQD/jz0A/447AP+LOgD/j00e/7KUfv/s6uf/////////////////////////////////////////////////39jT/5hnQ/+KOgD/jzwA/489AP+PPQD/jz0A/5A/Av+WShH/k0UL/489AP+URw7/lUgP/489AP+TRAn/lksS/5FABP+PPQD/jz0A/489AP+PPAD/izoA/5hnQ//f2NP//////////////////////////////////////+bj4P+TYDr/jDkA/489AP+PPQD/jz0A/489AP+OOwD/mVAa/9bDs//Bm37/l0wT/8qrlP/RuKX/mU8Z/7mLaf/byb3/oFwp/447AP+PPQD/jz0A/489AP+PPQD/jDkA/5NfOf/m4t/////////////////////////////7/Pz/qIhw/4o5AP+PPQD/jz0A/489AP+PPQD/jz0A/447AP+bVB//5N3X/+jj3//czcH/6ubj/+zp6P/dzsP/5d/Z/+rm5P+jYjL/jjsA/489AP+PPQD/jz0A/489AP+PPQD/ijkA/6aGbf/7+/v//////////////////////+Th3v+KTSD/jzwA/489AP+PPQD/jz0A/489AP+PPQD/jjsA/5tUH//j2tT/7/Dx/+/w8f/u7+//7u7v/+/w8f/v8PD/6eTh/6NiMf+OOwD/jz0A/489AP+PPQD/jz0A/489AP+PPAD/ikwd/+Pe2///////////////////////zcK6/4U7BP+OOgD/jToA/406AP+NOgD/jToA/406AP+MOAD/mlEc/+Pa1P/v7/D/7u7u/+7u7v/u7u7/7u7u/+7v7//p5OH/omAu/405AP+PPQD/jz0A/448AP+PPQD/jz0A/5A9AP+HPQb/y7+2///////////////////////OxsD/rYNj/7mLaf+5imj/uYpo/7mKaP+5imj/uYpo/7iJaP+/l3r/6OPg/+/v8P/u7+//7u/v/+7v7//u7+//7u/v/+zp5//Dn4P/vZN0/8qrlP/Kq5T/p2o8/448AP+PPQD/jz0A/4c7Av/Etar//////////////////////9bW1f/b19X/6ufk/+rm4//q5uP/6ubj/+rm4//q5uP/6ubj/+nm4//p5OH/6eTh/+nk4f/p5OH/6eTh/+nk4f/p5OH/6eTh/+nl4v/r6ef/8PHy//Hz9f/CnID/jjwA/489AP+PPQD/hzsC/8S1qv//////////////////////yr+3/5leMf+jYjL/o2Iy/6NiMv+jYTH/o2Iy/6NjM/+jYTD/pGAu/6RgLv+kYC7/pGAu/6RgLv+kYC7/pGAu/6RgLv+kYC7/pF8u/6hoOf+0fVb/s31V/5xRGv+POgD/kDsA/5A7AP+IOQD/xLWp///////////////////////Hua3/hjoA/447AP+OOwD/jjoA/4NTKP9nk5n/ZZmj/2yJh/9xfXL/cX1y/3F9cf9xfXH/cX1x/3F9cf9xfXH/cX1x/3F9cf9xfXL/cX1y/3B8cf9wfHH/cX50/3J/dv9yf3b/coB2/2x5cP+/wcD//////////////////////8e5rv+HPAP/jz0A/489AP+QOwD/eHJf/1DN/v9Pz///Usn5/1PG8v9Qzf//T8///0/P//9Pz///T8///0/P//9Pz///T87//1LI9f9UxfH/U8Xx/1PF8f9TxfH/U8Xx/1PF8f9TxvH/ULrj/7rP1v//////////////////////x7mu/4c8A/+PPQD/jz0A/488AP+LRxL/emxU/3hxXP9+YkP/fmRE/1e84P9Pzv//UM3//1DN//9Qzf//UM3//1DN//9Syfj/dndl/4RXMf+DWTH/g1kx/4NZMf+DWTH/g1kx/4NZMf98VTD/wrqz///////////////////////Hua7/hzwD/489AP+PPQD/jz0A/488AP+QOwD/kDoA/5A6AP+KSRb/WbjZ/0/O//9Qzf//UM3//1DN//9Qzf//UM3//1LI9v+AYT//kTkA/5A7AP+QOwD/kDsA/5A7AP+QOwD/kDsA/4g5AP/Etan//////////////////////8e5rv+HPAP/jz0A/489AP+PPQD/jz0A/489AP+PPQD/jzwA/4lLGf9ZuNn/T8///0/P//9Qzv//UM3//0/P//9Pzv//Usj2/39iQf+QOwD/jz0A/489AP+PPQD/jz0A/489AP+PPQD/hzsC/8S1qv//////////////////////x7mu/4c8A/+PPQD/jz0A/489AP+PPQD/jz0A/489AP+PPAD/iUsZ/1i73/9Yut3/a5CS/1i63v9TxvT/aZSZ/1+rwv9Ry/v/f2NC/5A7AP+PPQD/jz0A/489AP+PPQD/jz0A/489AP+HOwL/xLWq///////////////////////Hua7/hzwD/489AP+PPQD/jz0A/489AP+PPQD/jz0A/488AP+LRQ//boiF/3Z2ZP+PPQD/dndm/2+Hg/+MQwn/f2FA/2uRlP+FUyf/kDwA/489AP+PPQD/jz0A/489AP+PPQD/jz0A/4c7Av/Etar//////////////////////8a4rf+FOAD/jToA/486AP+QOwD/kDsA/5A7AP+QOwD/kDsA/5A6AP+QOgD/kDoA/5A7AP+QOgD/kDoA/5A6AP+QOgD/kDoA/5A6AP+QOwD/kDsA/5A7AP+QOwD/kDsA/486AP+NOgD/hTcA/8O0qf//////////////////////2dLM/6Z/Yv+pfl//emxX/2ZpW/9maVz/Z2pc/2ZrXv9ma1//Z2pd/2dpXP9maVz/Z2pc/2hsXP9obFz/Z2pc/2ZpXP9naVz/Z2pd/2ZrX/9ma17/Z2pc/2ZpXP9maVv/eWtX/6h+X/+mf2H/19DK/////////////////////////////v///9vf8v9Aac//O53f/1nC6v9NwfD/UKbF/1mQnv9MtNz/UMPy/1XE8P9Esun/L33Z/y982f9Dsen/VcTw/1HD8v9MtN3/WZCe/1Clw/9NwfD/WcLq/zuf3/8+ac7/2Nzx//7////////////////////////////////////i8Pb/gLzg/zqP1v9DqOD/VMHq/02+7f9RpMX/WpOg/0213P9SxPL/V8bx/0a27P8zg9z/M4Lc/0a16/9XxvH/UsTy/0223f9ak6H/UaPD/02+7f9Uwev/Q6ng/zmP1v99ut//4O/2///////////////////////7+vr/uNnm/2C23/9RY7n/dZLK/1p4w/9MbsP/SZDR/0qL0P9KjtT/S5TW/0mU1P9Jmtf/Sp/Z/0yV1P9MldT/Sp/Z/0ma1/9JlNT/S5TW/0qP1P9Ki9D/SZDR/0xvw/9ad8L/dZPK/1Fiuf9ftN//tdnm//v6+f//////+vv8/7/X4f9itNb/RKve/1Jkuv+EcLr/XnG+/0qBx/9Yqdb/SqLW/0xat/9kRq//jJrK/1tGsP9OWrj/SZbT/0mX0/9OXLj/WkWw/4yayv9lSK//TVm2/0qh1v9Yqdb/S4LH/11xvv+Ecbr/U2K5/0Sr3f9gs9b/vdbg//n7/P/F3+r/X8Dm/0iOzv9QUbX/Vjaw/1I8r/9Gk8//T67W/4vA1P9Oq9T/SJbS/1pLs/9xSrD/UHTE/0Sg1P9wt9X/crjV/0Sh1f9QdsT/cUuw/1tJsv9IldL/TavU/4vA1P9Rr9X/RpXP/1E+sP9WNa//UFC1/0iMzf9ewOb/w97p/9zg4f+iz+D/TZHN/2M4rf9cI6n/Vh+n/1Brv/9Rvuv/TbXk/0h9xv9SPa7/WCer/1gjqf9SR7L/RpHM/1e44P9YueD/RpPN/1JIsv9YI6n/WCer/1M8rv9Ie8X/TbXk/1G+6/9QbsD/Vh+n/1wjqf9kNqz/TY/N/5/O4P/b3+H/9fX0/8ze5f96wN7/f6/Y/52c0f+Xf8j/qJjK/7TW4/9wm8r/dFu4/2s4sP9dJqr/WSCp/1ccpv9MbcH/ZsTn/2fF6P9McML/Vx2m/1kgqf9dJqr/aziv/3VZuP9vmcn/s9bj/6may/+Xf8j/nZvR/4Cv1/95wN3/yt3k//X09P/+/v7/8O/u/+Hm6P+v0t//nc7f/5jP5P+n0uL/rM3a/6/K1v+av9X/rLzX/5KWyv+IeMH/iGa+/3F7xf+qzNr/rM3a/3F9xv+IZr//iHjC/5KVyf+rvNf/mr/V/67J1v+szdr/p9Li/5jP5P+czt//rtLf/+Dm6P/v7+7//f39/////////////f39//Ty8f/x8vL/4+rt/9bg5P/V5On/z9ne/8bU2v+5zdP/uM7U/6TDzv+zy9X/o8bW/5m8yv+bvMr/osXW/7TL1v+jw87/uc7V/7jN0//G1dr/ztne/9Xj6f/W4OT/4urt//Hy8v/08vH//f38//////////////////////////////////////////7//v39///////7+/r//Pv7//X09P/08/P/6erq/+Pl5v/Eysv/kpaQ/5KWkP/Dycn/4+Xm/+nq6v/08/P/9fT0//z7+//7+/r////+//79/f////7//////////////////////////////////////////////////////////////////////////////////////////////////////+bk4f+GlpT/hJaU/+Ti3///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+vv7/7fT3f+20tz/+vv7////////////////////////////////////////////////////////////////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=
"@
   
    # Convert Base64 string to byte array and create the icon
    $iconBytes = [Convert]::FromBase64String($iconBase64)
    $iconStream = New-Object System.IO.MemoryStream(,$iconBytes)
    $form.Icon = New-Object System.Drawing.Icon($iconStream)

    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Size = New-Object System.Drawing.Size(460, 320)  # Adjusted size to fit within the form
    $richTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $richTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $richTextBox.ScrollBars = "Vertical"  # Add vertical scrollbar
    $form.Controls.Add($richTextBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(200, 340)  # Adjusted position to fit within the form
    $okButton.Add_Click({ $form.Close() })
    $form.Controls.Add($okButton)

    return $form, $richTextBox
}

# Main script
$form, $richTextBox = Create-Form

# ... (rest of your script)

function Add-Text {
    param ($richTextBox, $text, $color)
    $richTextBox.SelectionColor = [System.Drawing.Color]::FromName($color)
    $richTextBox.AppendText($text + "`n")
    $richTextBox.SelectionColor = $richTextBox.ForeColor
}

function Add-HorizontalLine {
    param ($richTextBox)
    $richTextBox.AppendText("`n" + ("-" * 50) + "`n")  # Add a horizontal line
}

function Check-TPM {
    $tpmState = Get-Tpm | Select-Object TpmReady
    if ($tpmState.TpmReady -eq $false) {
        return [PSCustomObject]@{ Text = "TPM är redo för användning: Nej"; Color = "Red" }
    } else {
        return [PSCustomObject]@{ Text = "TPM är redo för användning: Ja"; Color = "Green" }
    }
}

function Check-Certificate {
    Set-Location cert:\LocalMachine\My
    $LMCertExists = $false
    $LMCertSubject = "NONE"
    $LMCertStore = Get-ChildItem
    foreach($certificate in $LMCertStore){
        if($certificate.Subject -like "*$env:COMPUTERNAME*"){
            $LMCertExists = $true
            $LMCertSubject = $certificate.Subject
        }
    }
    Set-Location C:
    if ($LMCertExists) {
        return [PSCustomObject]@{ Text = "Datorns enhetscertifikat: Ja ($LMCertSubject)"; Color = "Green" }
    } else {
        return [PSCustomObject]@{ Text = "Datorns enhetscertifikat: Nej"; Color = "Red" }
    }
}

function Check-SecureBoot {
    $dgSecureBootEnabled = Confirm-SecureBootUEFI
    if ($dgSecureBootEnabled -eq "True") {
        return [PSCustomObject]@{ Text = "UEFI Secure Boot är aktiverat: Ja"; Color = "Green" }
    } else {
        return [PSCustomObject]@{ Text = "UEFI Secure Boot är aktiverat: Nej"; Color = "Red" }
    }
}

function Check-Application {
    param ($appName, $displayName)
    $installedApps = Get-WmiObject -Class Win32_Product
    $installedApp = $installedApps | Where-Object Name -Like "*$appName*"
    if ($installedApp.Name.Count -gt 0) {
        return [PSCustomObject]@{ Text = "$($displayName): Ja ($($installedApp.Name.Count))"; Color = "Green" }
    } else {
        return [PSCustomObject]@{ Text = "$($displayName): Nej"; Color = "Red" }
    }
}

function Set-ComplianceWarning {
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    $registryName = "legalnoticetext"
    $registryValue = "Datorns installation har inte slutförts korrekt och bör inte användas förrän detta åtgärdats. Kontakta servicedesk för åtgärd."
    New-ItemProperty -Path $registryPath -Name $registryName -Value $registryValue -Force | Out-Null
    $registryName = "legalnoticecaption"
    $registryValue = "VARNING"
    New-ItemProperty -Path $registryPath -Name $registryName -Value $registryValue -Force | Out-Null
}

# Main script
$form, $richTextBox = Create-Form

Add-Text -richTextBox $richTextBox -text "Efterkontroll Operativsystem SMP typ 2" -color "Black"
Add-Text -richTextBox $richTextBox -text "Denna rapport genereras under OSD Task Sequence." -color "Black"
Add-Text -richTextBox $richTextBox -text "Om fel detekteras visas rapporten på enhetens skärm." -color "Black"
Add-HorizontalLine -richTextBox $richTextBox  # Add horizontal line

$isCompliant = $true

$tpmResult = Check-TPM
Add-Text -richTextBox $richTextBox -text $tpmResult.Text -color $tpmResult.Color
$isCompliant = $isCompliant -and ($tpmResult.Color -eq "Green")


$certResult = Check-Certificate
Add-Text -richTextBox $richTextBox -text $certResult.Text -color $certResult.Color
$isCompliant = $isCompliant -and ($certResult.Color -eq "Green")


$secureBootResult = Check-SecureBoot
Add-Text -richTextBox $richTextBox -text $secureBootResult.Text -color $secureBootResult.Color
$isCompliant = $isCompliant -and ($secureBootResult.Color -eq "Green")


$trellixResult = Check-Application -appName "Trellix Agent" -displayName "Trellix Agenter installerade"
Add-Text -richTextBox $richTextBox -text $trellixResult.Text -color $trellixResult.Color
$isCompliant = $isCompliant -and ($trellixResult.Color -eq "Green")


$mbamResult = Check-Application -appName "MDOP MBAM" -displayName "MDOP MBAM Agenter installerade"
Add-Text -richTextBox $richTextBox -text $mbamResult.Text -color $mbamResult.Color
$isCompliant = $isCompliant -and ($mbamResult.Color -eq "Green")


$officeResult = Check-Application -appName "Office" -displayName "Office program installerade"
Add-Text -richTextBox $richTextBox -text $officeResult.Text -color $officeResult.Color
$isCompliant = $isCompliant -and ($officeResult.Color -eq "Green")

if ($isCompliant -eq $false) {
    Set-ComplianceWarning
}

$form.ShowDialog()