Import-Csv .\resultado.csv | ForEach-Object {
    $userInfo = "$($_.nome), $($_.dn), $($_.primeironome), $($_.Sobrenome), $($_.conta), $($_.email), $($_.Desc), $($_.Office), $($_.Dep), $($_.ou), $($_.pass)"
    Add-Content -Path "usuarios_simulados.csv" -Value $userInfo
}
