param(
    [string]$FirstArgument
)

if ($FirstArgument -ne "VALID") {
    $process = Start-Process -FilePath "cmd" -ArgumentList "/c", "install.bat" -Wait -PassThru -NoNewWindow
    exit $process.ExitCode
}
# ----------------------------------------------------------------------------------------------

# Controllo versione .NET Framework 4.7.2
Write-Host "[..] Verifica versione .NET Framework 4.7.2..."

try {
    $release = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release
    
    if ($release -ge 461808) {
        if ($release -ge 528040) {
            Write-Host "[OK] .NET Framework 4.8 o superiore installato (release: $release)"
        } elseif ($release -ge 461814) {
            Write-Host "[OK] .NET Framework 4.7.2 installato (release: $release)"
        } else {
            Write-Host "[OK] .NET Framework 4.7.2 installato (release: $release)"
        }
        
        Write-Host "[OK] Versione .NET Framework 4.7.2 o superiore confermata"
        Write-Host ""
        
    } else {
        Write-Host "[KO] .NET Framework 4.7.2 non installato"
        Write-Host "[KO] Release trovato: $release"
        Write-Host "[KO] Release minimo richiesto: 461808"
        Write-Host "[KO] Installazione terminata per requisiti non soddisfatti"
        exit 1
    }
    
} catch {
    Write-Host "[KO] Errore nel leggere il registro .NET Framework"
    Write-Host "[KO] $_"
    Write-Host "[KO] Installazione terminata per errore di sistema"
    exit 1
}

# ----------------------------------------------------------------------------------------------

# Controllo e installazione modulo MilestonePSTools
Write-Host "[..] Verifica modulo PowerShell MilestonePSTools..."

try {
    $moduleName = "MilestonePSTools"
    $moduleInstalled = Get-Module -ListAvailable -Name $moduleName
    
    if ($moduleInstalled) {
        Write-Host "[OK] Modulo $moduleName installato"
        Write-Host "[??] Versione: $($moduleInstalled.Version)"
    } else {
        Write-Host "[!!] Modulo $moduleName non trovato"
        Write-Host "[..] Installazione modulo $moduleName a livello utente..."
        
        # Percorso moduli utente
        $userModulesPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"
        $moduleDestPath = "$userModulesPath\$moduleName"
        
        # Crea directory se non esiste
        if (!(Test-Path $userModulesPath)) {
            New-Item -ItemType Directory -Path $userModulesPath -Force | Out-Null
        }
        
        # Cerca il modulo nella directory corrente del progetto
        $projectModulePath = ".\$moduleName"
        if (Test-Path $projectModulePath) {
            Write-Host "[..] Copia modulo da $projectModulePath a $moduleDestPath"
            Copy-Item -Path $projectModulePath -Destination $moduleDestPath -Recurse -Force
            
            # Estrazione file avcodec-61.dll.zip dalla cartella bin del modulo
            $zipFilePath = "$moduleDestPath\bin\avcodec-61.dll.zip"
            $binFolderPath = "$moduleDestPath\bin"
            
            if (Test-Path $zipFilePath) {
                Write-Host "[..] Estrazione file avcodec-61.dll.zip..."
                try {
                    # Estrae il file zip nella stessa cartella bin
                    Expand-Archive -Path $zipFilePath -DestinationPath $binFolderPath -Force
                    Write-Host "[OK] File avcodec-61.dll estratto con successo"
                    
                    # Elimina il file zip dopo l'estrazione
                    Remove-Item -Path $zipFilePath -Force
                    Write-Host "[OK] File zip eliminato: avcodec-61.dll.zip"
                    
                } catch {
                    Write-Host "[KO] Errore nell'estrazione del file zip: $_"
                    exit 1
                }
            } else {
                Write-Host "[!!] File avcodec-61.dll.zip non trovato in $zipFilePath"
                Write-Host "[!!] Continuando senza estrazione..."
            }
            
            # Verifica installazione
            $moduleCheck = Get-Module -ListAvailable -Name $moduleName
            if ($moduleCheck) {
                Write-Host "[OK] Modulo $moduleName installato con successo"
                Write-Host "[??] Installato in: $moduleDestPath"
            } else {
                Write-Host "[KO] Errore nell'installazione del modulo $moduleName"
                exit 1
            }
        } else {
            Write-Host "[KO] Modulo $moduleName non trovato nella directory del progetto"
            Write-Host "[KO] Percorso cercato: $projectModulePath"
            exit 1
        }
    }
    
} catch {
    Write-Host "[KO] Errore durante la gestione del modulo $moduleName"
    Write-Host "[KO] $_"
    exit 1
}

Write-Host ""

# Esegue il comando ls se tutto è OK
Write-Host "[..] Esecuzione comando ls..."
Get-ChildItem

Write-Host ""
Write-Host "[OK] Script completato con successo"
pause

