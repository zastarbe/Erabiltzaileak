# -Erabiltzaileak sarean gehitu
# -Erabiltzaileen karpeta egin
# -Taldeen karpeta egin

######################

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

######################


$inputfile = Get-FileName "R:\Proyectos ABIERTOS\201809 erabiltzaileak\ikasleak"
$inputdata = get-content $inputfile


# Aldagaiak
$NunSartu = "OU=Alumnos,OU=Usuarios,OU=CEINPRO,DC=ceinpro,DC=es"
$LanekoPath = "C:\BD"
$ErabiltzaileFile = $inputfile
$LogFile = ("AD-Erabiltzaileak-{0:yyyy-MM-dd-HH-mm-ss}.log" -f (Get-Date)) 
$Log = "$LanekoPath\$LogFile" 
$LogonScript = "alumno.bat"

# Log hasiera
Start-Transcript $Log 



# Erabiltzaileak sortu
#$Users = Import-Csv -Path "$LanekoPath\$ErabiltzaileFile"            
$Users = Import-Csv -Path "$ErabiltzaileFile" 
foreach ($User in $Users)            
{            
    # Erabiltzailearen aldagaiak
    $Displayname = $User.'Nombre' + " " + $User.'Apellidos'            
    $ErabIzena = $User.'Nombre'            
    $ErabAbizena = $User.'Apellidos'            
    $OU = $NunSartu
    $SAM = $User.'Usuario'            
    $UPN = $User.'Usuario' + "@ceinpro.es"
    $Korreo = $User.'Usuario' + "@alumni.ceinpro.es"
    $Taldea = $User.'Grupo'
    $IkasleTaldea = "Alumnos" + $User.'Grupo'
    $IrakasleTaldea = "Profesores" + $User.'Grupo'
#    $Password = "Ceinpro1."            #Lehen hau
    $Password = $User.'Password'       #Orain hau


    # TALDEAREN KARPETA Aldagaiak
    $Zerbitzaria = "Heriz16"
    $Diska = "E"
    $TaldePath = "Ikasleak\Grp\$Taldea"
    $TaldePathOsoa = "\\$Zerbitzaria\$Diska$\$TaldePath"
    $IkasleTaldea = "Alumnos$Taldea"
    $IrakasleTaldea = "Profesores$Taldea"

    # ERABILTZAILEAREN KARPETA Aldagaiak
    $IkaslePath = "Ikasleak\Dok\$SAM"
    $IkaslePathOsoa = "\\$Zerbitzaria\$Diska$\$IkaslePath"




    # Tilera Pasa Izenak
    $ErabIzenaTile = (Get-Culture).TextInfo.ToTitleCase($ErabIzena.ToLower())
    $ErabAbizenaTile = (Get-Culture).TextInfo.ToTitleCase($ErabAbizena.ToLower())
    $IkusIzenaTile = $ErabIzenaTile + " " + $ErabAbizenaTile


    Write-Host "-----------------------------------"
    Write-Host "Erabiltzailea: $IkusizenaTile - $SAM"



    # Erabiltzailea sortu
    Try
    {
        New-ADUser -Name "$IkusIzenaTile" -DisplayName "$IkusIzenaTile" -SamAccountName $SAM -UserPrincipalName $UPN -GivenName "$ErabIzenaTile" -Surname "$ErabAbizenaTile" -EmailAddress $Korreo -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled 1 -Path $OU -ChangePasswordAtLogon 1 -ScriptPath $LogonScript 
        Write-Host "Erabiltzaile berria: $IkusizenaTile - $SAM"
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        Write-Host "$ErrorMessage - $SAM"
        Write-Host "---"
        Write-Host "Erabiltzaile bazegoen: $IkusizenaTile - $SAM"
        Write-Host "---"
        Write-Host "$SAM erabiltzailea tokiz mugitzen"
        $erabil = Get-ADObject -Filter "sAMAccountName -eq '$SAM'" -SearchBase 'OU=Usuarios,OU=CEINPRO,DC=ceinpro,DC=es' | foreach {$_.DistinguishedName}
        Write-Host "Lehen: $erabil"
        Move-ADObject -Identity $erabil -TargetPath "OU=Alumnos,OU=Usuarios,OU=CEINPRO,DC=ceinpro,DC=es"
        $erabil = Get-ADObject -Filter "sAMAccountName -eq '$SAM'" -SearchBase 'OU=Usuarios,OU=CEINPRO,DC=ceinpro,DC=es' | foreach {$_.DistinguishedName}
        Write-Host "Orain: $erabil"
        Write-Host "---"
        Write-Host "Pasahitza berria: $Password"
        #$passberria = $Password As SecureString
        $passberria = ConvertTo-SecureString -String "$Password" -AsPlainText –Force
        Set-ADAccountPassword $SAM -NewPassword $passberria –Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True -Enabled $True
        #Set-ADAccountPassword $SAM -NewPassword $passberria -Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True -Enabled 1
        #Set-ADUser -SamAccountName $SAM -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled 1 -Path $OU -ChangePasswordAtLogon 1 -ScriptPath $LogonScript 
        Write-Host "---"
        #Break 
    }


    # Erabiltzailea taldean sartu
    Try
    {
 	    Add-ADGroupMember -Identity $IkasleTaldea -Members $SAM
        Write-Host "Erabiltzailea talde honetan: $IkasleTaldea"
        Write-Host "---"
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erabiltzailea talde honetan bazegoen: $IkasleTaldea"
        Write-Host "$ErrorMessage - $SAM"
        Write-Host "---"
        Break
    }


    # ERABILTZAILEAREN KARPETA

    #Ikaslearen karpeta existitzen da?
    if (Test-Path $IkaslePathOsoa) {
    Write-Host "Ikaslearen karpeta existitzen da."
    Write-Host "---"
    } 
    else 
    {
    #Ikaslearen karpeta egin eta eskubideak gehitu
    New-Item -Path $IkaslePathOsoa -type directory -Force 
    Add-NTFSAccess –Path $IkaslePathOsoa –Account $SAM –AccessRights Modify
    Write-Host "---"
    }




    # TALDEAREN KARPETA

    #Taldearen karpeta existitzen da?
    if (Test-Path $TaldePathOsoa) {
    Write-Host "Taldearen karpeta existitzen da."
    Write-Host "---"
    } 
    else 
    {
    # Taldearen karpeta egin eta eskubideak gehitu
    New-Item -Path $TaldePathOsoa -type directory -Force 
    Add-NTFSAccess –Path $TaldePathOsoa –Account $IkasleTaldea –AccessRights ReadAndExecute
    Add-NTFSAccess –Path $TaldePathOsoa –Account $IrakasleTaldea –AccessRights Modify
    # Irakasle karpeta
    New-Item -Path $TaldePathOsoa\Profesor -type directory -Force 
    Add-NTFSAccess –Path $TaldePathOsoa\Profesor –Account $IkasleTaldea –AccessRights Write
    Add-NTFSAccess -Path $TaldePathOsoa\Profesor -Account $IkasleTaldea -AccessType Deny -AccessRights Read
    Write-Host "---"
    }

}

#Send-MailMessage -From erabiltzaileak@ceinpro.es -To informatika@ceinpro.es -Subject "Erabiltzaile berriak!" -SmtpServer 192.168.30.9 -Body "Erabiltzaile berriak egin dira: + $ErabIzena"
