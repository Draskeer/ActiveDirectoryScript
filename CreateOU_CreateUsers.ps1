Import-Module ActiveDirectory

#Crée un utilisateur dans l'Organizational Unit (OU) correspondant à son pays et à son poste. Ajoute la création de l'utilisateur au fichier de log.
function createUser(){
    param(
        [int]$id,
        [string]$name,
        [string]$country,
        [string]$position
    )
    try {
        $userName = "${id}_${name}"
        $password = "Pa`$`$w0rd" | ConvertTo-SecureString -AsPlainText -Force
        $ouPath = "OU=$position,OU=$country,OU=TPFinal,DC=ASI,DC=local"
        New-ADUser -SamAccountName $userName -UserPrincipalName "$userName@ASI.Local" -Name $userName -Enabled $true -AccountPassword $password -Path $ouPath
        createLogs -userName $userName -id $id -ouPath $ouPath
    }catch {
        $errorMessage = "An error occurred: $_"
        $errorMessage
        $errorFilePath = Join-Path -Path $PSScriptRoot -ChildPath "error.txt"
        $errorMessage | Out-File -FilePath $errorFilePath -Append
    }

}

#Crée les différents Organizational Units (OU) en fonction des postes et des pays des utilisateurs du fichier CSV.
function createOU {
    try {
        $csvFile = "./users.csv"
        $data = Import-Csv -Path $csvFile -Delimiter '|'
        $countryList = $data | Select-Object -ExpandProperty country -Unique
        $positionList = $data | Select-Object -ExpandProperty position -Unique
        foreach ($country in $countryList) {
            $ouPath = "OU=TPFinal,DC=ASI,DC=local"
            $countryOU = New-ADOrganizationalUnit -Name $country -Path $ouPath -ProtectedFromAccidentalDeletion $False
            foreach ($position in $positionList) {
                $ouPath = "OU=$country,OU=TPFinal,DC=ASI,DC=local"
                New-ADOrganizationalUnit -Name $position -Path $ouPath -ProtectedFromAccidentalDeletion $False
            }       
        }
    }
    catch {
        $errorMessage = "An error occurred: $_"
        $errorMessage
        $errorFilePath = Join-Path -Path $PSScriptRoot -ChildPath "error.txt"
        $errorMessage | Out-File -FilePath $errorFilePath -Append
    }
}

#Parcourt le fichier CSV afin de pouvoir récupérer les informations des utilisateurs pour appeler la fonction qui va les créer.
function importUsers {
    try {
        $csvFile = "./users.csv"
        $data = Import-Csv -Path $csvFile -Delimiter '|'
        foreach ($raw in $data) {
            $id = $raw.id
            $name = $raw.name
            $position = $raw.position
            $country = $raw.country
            createUser -id $id -name $name -country $country -position $position
        }  
    }catch{
        $errorMessage = "An error occurred: $_"
        $errorMessage
        $errorFilePath = Join-Path -Path $PSScriptRoot -ChildPath "error.txt"
        $errorMessage | Out-File -FilePath $errorFilePath -Append
    }
}

#Analyse le fichier CSV afin de récupérer des listes et d'en extraire des informations utiles.
function analyseData {
    try {
        " "
        $csvFile = "./users.csv"
        $data = Import-Csv -Path $csvFile -Delimiter '|'
        "il y a $($data.count) utilisateurs sur le domaine."
        $countryList = $data | Select-Object -ExpandProperty country -Unique
        " "
        "============== Il y a $($countryList.count) pays. ======================"
        " "
        foreach ($country in $countryList) {
            $userList = $data | Where-Object { $_.country -eq $country }
            "il y a $($userList.count) utilisateur en $country."
        }
        $positionList = $data | Select-Object -ExpandProperty position -Unique
        " "
        "============== Il y a $($positionList.count) positions. ================="
        " "
        foreach ($position in $positionList) {
            $userList = $data | Where-Object { $_.position -eq $position }
            "il y a $($userList.count) utilisateur occupant le poste de $position."
        }
        " "
    }catch{
        $errorMessage = "An error occurred: $_"
        $errorMessage
        $errorFilePath = Join-Path -Path $PSScriptRoot -ChildPath "error.txt"
        $errorMessage | Out-File -FilePath $errorFilePath -Append
    }
}

#Permet de créer des lignes à ajouter au fichier log.txt.
function createLogs {
    param(
        $userName,
        $id,
        $ouPath
    )
    try {
        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "Utilisateur : $userName, ID : $id, OU de destination : $ouPath, Date de création : $date"
        $logEntry | Out-File -Append -FilePath "./log.txt"
    }catch{
        $errorMessage = "An error occurred: $_"
        $errorMessage
        $errorFilePath = Join-Path -Path $PSScriptRoot -ChildPath "error.txt"
        $errorMessage | Out-File -FilePath $errorFilePath -Append
    }
}

#Fonction principale du programme qui permet d'appeler les différentes fonctions souhaitées.
function ADManager {
    "=========== Liste des options disponible : =============="
    "1- Import Users"
    "2- Analyse Data"
    "3- Create OU"
    $i = Read-Host "Entre l'index de l'option que vous voulez: "
    switch ($i) {
        "1" { importUsers }
        "2" { analyseData }
        "3" { createOU }
        Default { 
            "Veuillez entrer une valeur existante."
        }
    }
}

ADManager