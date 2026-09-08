param([Parameter(Mandatory=$true)][string]$PastaPublicacao)
$ErrorActionPreference = 'Stop'
$identidade = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identidade)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Execute este script em um PowerShell como administrador.'
}
$nome = 'PsMonitorArquivos'
if (Get-Service -Name $nome -ErrorAction SilentlyContinue) {
    throw 'O servico PsMonitorArquivos ja existe. Confira a instalacao existente antes de publicar outra versao.'
}
$pasta = (Resolve-Path -LiteralPath $PastaPublicacao).Path
$exe = Join-Path $pasta 'MonitorArquivos.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw 'MonitorArquivos.exe nao encontrado. Execute dotnet publish primeiro.' }
New-Service -Name $nome -DisplayName $nome -BinaryPathName ('"' + $exe + '"') -StartupType Manual | Out-Null
Start-Service -Name $nome
Get-Service -Name $nome
