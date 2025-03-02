# Importar o módulo do Active Directory
Import-Module ActiveDirectory

# Definir a OU de origem (onde os usuários estão atualmente)
$ouOrigem = "OU=Usuarios,OU=CURSOS,OU=SENAINMT,OU=SENAI,DC=SFIEMT-EDU,DC=teste,DC=local"

# Definir a OU de destino (para onde os usuários serão movidos)
$ouDestino = "OU=NovosUsuarios,OU=CURSOS,OU=SENAINMT,OU=SENAI,DC=SFIEMT-EDU,DC=teste,DC=local"

# Verificar se as OUs existem
try {
    $ouOrigemObj = Get-ADOrganizationalUnit -Identity $ouOrigem -ErrorAction Stop
    Write-Host "OU de origem encontrada: $($ouOrigemObj.DistinguishedName)" -ForegroundColor Green
} catch {
    Write-Host "Erro: OU de origem não encontrada. Verifique o Distinguished Name (DN)." -ForegroundColor Red
    exit
}

try {
    $ouDestinoObj = Get-ADOrganizationalUnit -Identity $ouDestino -ErrorAction Stop
    Write-Host "OU de destino encontrada: $($ouDestinoObj.DistinguishedName)" -ForegroundColor Green
} catch {
    Write-Host "Erro: OU de destino não encontrada. Verifique o Distinguished Name (DN)." -ForegroundColor Red
    exit
}

# Buscar todos os usuários na OU de origem
$usuarios = Get-ADUser -SearchBase $ouOrigem -Filter *

# Verificar se há usuários na OU de origem
if ($usuarios.Count -eq 0) {
    Write-Host "Nenhum usuário encontrado na OU de origem." -ForegroundColor Yellow
    exit
}

# Mover todos os usuários para a OU de destino
$usuarios | ForEach-Object {
    Move-ADObject -Identity $_ -TargetPath $ouDestino
    Write-Host "Usuário $($_.Name) movido para $ouDestino" -ForegroundColor Green
}

Write-Host "Operação concluída." -ForegroundColor Green