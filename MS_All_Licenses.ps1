#####################################################
## Jornada365 | Wintuner                         ####
## jornada365.cloud                              ####
## Sua jornada comeca aqui.                      ####
#####################################################

# Configuracao de tratamento de erros e modo rigoroso
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Importacao dos modulos necessarios
function Import-RequiredModules {
    try {
        Write-LogMessage -Message "Verificando modulos necessarios..." -Type INFO
        $requiredModules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Groups",
            "Microsoft.Graph.Identity.DirectoryManagement"
        )
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -Name $module -ListAvailable)) {
                Write-LogMessage -Message "Instalando modulo $module..." -Type INFO
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
            }
            Import-Module -Name $module -Force
        }
        Write-LogMessage -Message "Todos os modulos necessarios estao disponiveis." -Type SUCCESS
    } catch {
        Write-LogMessage -Message "Erro ao importar modulos: $_" -Type ERROR
        throw $_
    }
}

# Funcao para exibir mensagens formatadas
function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Type
    )

    $colors = @{ 'INFO' = 'White'; 'SUCCESS' = 'Green'; 'WARNING' = 'Yellow'; 'ERROR' = 'Red'; 'DEBUG' = 'Cyan' }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp][$Type] $Message" -ForegroundColor $colors[$Type]
}

# Funcao para autenticar no Microsoft Graph
function Connect-GraphPersistent {
    param(
        [switch]$Reconnect
    )
    if (-not (Get-MgContext) -or $Reconnect) {
        try {
            Write-LogMessage -Message "Conectando ao Microsoft Graph..." -Type INFO
            Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
            Write-LogMessage -Message "Autenticado com sucesso no Microsoft Graph." -Type SUCCESS
        } catch {
            Write-LogMessage -Message "Erro ao tentar autenticar no Microsoft Graph: $_" -Type ERROR
            throw $_
        }
    } else {
        Write-LogMessage -Message "Sessao do Microsoft Graph ja esta ativa." -Type INFO
    }
}

# Funcao para listar licencas (Sku IDs)
function Listar-Licencas {
    try {
        Write-LogMessage -Message "Obtendo licencas disponiveis..." -Type INFO
        $licenses = Get-MgSubscribedSku -All | Sort-Object SkuPartNumber
        if (-not $licenses) {
            Write-LogMessage -Message "Nenhuma licenca encontrada no tenant." -Type ERROR
            return $null
        }

        $licenseMapping = @{ }
        $index = 1
        foreach ($license in $licenses) {
            if ($license.SkuPartNumber) {
                $licenseMapping[$index] = $license
                Write-Host "" -ForegroundColor Blue
                Write-Host "$index - $($license.SkuPartNumber)" -ForegroundColor Blue
                Write-Host "Total: $($license.PrepaidUnits.Enabled)" -ForegroundColor White
                Write-Host "Em Uso: $($license.ConsumedUnits)" -ForegroundColor White
                Write-Host "Disponiveis: $(($license.PrepaidUnits.Enabled) - $($license.ConsumedUnits))" -ForegroundColor White
                Write-Host "=====================" -ForegroundColor Blue
                Write-Host "" -ForegroundColor Blue
                $index++
            }
        }
        return $licenseMapping
    } catch {
        Write-LogMessage -Message "Erro ao listar licencas: $_" -Type ERROR
        return $null
    }
}

# Funcao para gerenciar membros do grupo com base no Sku ID selecionado
function Gerenciar-UsuariosDoGrupo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SkuId,

        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    Write-LogMessage -Message "Iniciando gestao dos usuarios do grupo..." -Type INFO
    try {
        # Inicializar contadores
        $addedCount = 0
        $removedCount = 0

        # Obter o grupo
        $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        Write-LogMessage -Message "Grupo encontrado: $($group.DisplayName)" -Type INFO

        # Obter todos os usuarios licenciados
        Write-LogMessage -Message "Obtendo usuarios com a licenca selecionada..." -Type INFO
        $usersWithLicense = @(Get-MgUser -All -Filter "assignedLicenses/any(x:x/skuId eq $SkuId)" -Select "Id,DisplayName,UserPrincipalName")

        # Obter membros atuais do grupo
        Write-LogMessage -Message "Obtendo membros atuais do grupo..." -Type INFO
        $currentMembers = @(Get-MgGroupMember -GroupId $GroupId -All | Select-Object Id, AdditionalProperties)

        # Identificar membros para adicionar e remover
        $usersWithLicenseIds = @($usersWithLicense | Select-Object -ExpandProperty Id)
        $currentMemberIds = @($currentMembers | Select-Object -ExpandProperty Id)

        # Usuarios para adicionar (tem licenca mas nao estao no grupo)
        $usersToAdd = @($usersWithLicense | Where-Object { $_.Id -notin $currentMemberIds })

        # Usuarios para remover (estao no grupo mas nao tem licenca)
        $usersToRemove = @($currentMembers | Where-Object { $_.Id -notin $usersWithLicenseIds })

        # Adicionar usuarios ao grupo
        foreach ($user in $usersToAdd) {
            try {
                New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id
                Write-LogMessage -Message "Usuario '$($user.DisplayName)' ($($user.UserPrincipalName)) adicionado ao grupo." -Type SUCCESS
                $addedCount++
            } catch {
                Write-LogMessage -Message "Erro ao adicionar o usuario '$($user.DisplayName)' ao grupo: $_" -Type ERROR
            }
        }

        # Remover usuarios do grupo
        foreach ($user in $usersToRemove) {
            try {
                Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id
                $userName = (Get-MgUser -UserId $user.Id).DisplayName
                Write-LogMessage -Message "Usuario '$userName' removido do grupo." -Type SUCCESS
                $removedCount++
            } catch {
                Write-LogMessage -Message "Erro ao remover o usuario com ID '$($user.Id)' do grupo: $_" -Type ERROR
            }
        }

        Write-LogMessage -Message "Gestao dos usuarios do grupo concluida." -Type SUCCESS

        # Exibir resumo
        Write-LogMessage -Message "Resumo da operacao:" -Type INFO
        Write-LogMessage -Message "Usuarios adicionados: $addedCount" -Type INFO
        Write-LogMessage -Message "Usuarios removidos: $removedCount" -Type INFO

        # **Ajuste realizado aqui**
        $groupMembers = @(Get-MgGroupMember -GroupId $GroupId -All)
        $totalMembers = $groupMembers.Count
        Write-LogMessage -Message "Total de membros atual: $totalMembers" -Type INFO

    } catch {
        Write-LogMessage -Message "Erro ao gerenciar os usuarios do grupo: $_" -Type ERROR
    }
}

# Funcao principal para execucao do script
function Main {
    try {
        # Importar modulos necessarios
        Import-RequiredModules
        
        # Conectar ao Graph
        Connect-GraphPersistent

        do {
            $licenseMapping = Listar-Licencas
            if (-not $licenseMapping) {
                return
            }

            # Selecionar a licenca
            do {
                $licenseChoice = Read-Host "Digite o numero da licenca que deseja processar"
                $isValidLicense = [int]::TryParse($licenseChoice, [ref]$null) -and $licenseMapping.ContainsKey([int]$licenseChoice)
                if (-not $isValidLicense) {
                    Write-LogMessage -Message "Escolha invalida para licenca. Tente novamente." -Type ERROR
                }
            } while (-not $isValidLicense)
            
            $selectedLicense = $licenseMapping[[int]$licenseChoice]
            Write-LogMessage -Message "Licenca selecionada: $($selectedLicense.SkuPartNumber) | Sku ID: $($selectedLicense.SkuId)" -Type INFO

            # Solicitar o ID do grupo
            $groupId = Read-Host "Digite o ID do grupo que deseja atualizar"

            # Gerenciar a associacao do grupo
            Gerenciar-UsuariosDoGrupo -SkuId $selectedLicense.SkuId -GroupId $groupId

            # Perguntar ao usuario se deseja continuar ou sair
            $continuar = Read-Host "Deseja continuar gerenciando outro grupo ou licenca? (Sim/Nao)"
        } while ($continuar -match '^(S|s)(im)?$')
    } catch {
        Write-LogMessage -Message "Erro critico na execucao do script: $_" -Type ERROR
    } finally {
        # Desconectar do Microsoft Graph se uma sessao for iniciada
        if (Get-MgContext) {
            Disconnect-MgGraph -Confirm:$false
        }
        Write-LogMessage -Message "Execucao do script finalizada" -Type INFO
    }
}

# Execucao da funcao principal
Main
