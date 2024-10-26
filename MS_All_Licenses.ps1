#####################################################
## Jornada365 | Wintuner                         ####
## jornada365.cloud                              ####
## Sua jornada começa aqui.                      ####
## ##########################################    ####
#####################################################

# Configuração de tratamento de erros e modo rigoroso
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Importação dos módulos necessários
function Import-RequiredModules {
    try {
        Write-LogMessage -Message "Verificando módulos necessários..." -Type INFO
        $requiredModules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Groups",
            "Microsoft.Graph.Identity.DirectoryManagement"
        )
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -Name $module -ListAvailable)) {
                Write-LogMessage -Message "Instalando módulo $module..." -Type INFO
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
            }
            Import-Module -Name $module -Force
        }
        Write-LogMessage -Message "Todos os módulos necessários estão disponíveis." -Type SUCCESS
    } catch {
        Write-LogMessage -Message "Erro ao importar módulos: $_" -Type ERROR
        throw $_
    }
}

# Função para exibir mensagens formatadas
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

# Função para autenticar no Microsoft Graph
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
        Write-LogMessage -Message "Sessão do Microsoft Graph já está ativa." -Type INFO
    }
}

# Função para listar licenças (Sku IDs)
function Listar-Licencas {
    try {
        Write-LogMessage -Message "Obtendo licenças disponíveis..." -Type INFO
        $licenses = Get-MgSubscribedSku -All | Sort-Object SkuPartNumber
        if (-not $licenses) {
            Write-LogMessage -Message "Nenhuma licença encontrada no tenant." -Type ERROR
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
                Write-Host "Disponíveis: $(($license.PrepaidUnits.Enabled) - $($license.ConsumedUnits))" -ForegroundColor White
                Write-Host "=====================" -ForegroundColor Blue
                Write-Host "" -ForegroundColor Blue
                $index++
            }
        }
        return $licenseMapping
    } catch {
        Write-LogMessage -Message "Erro ao listar licenças: $_" -Type ERROR
        return $null
    }
}

# Função para gerenciar membros do grupo com base no Sku ID selecionado
function Gerenciar-UsuariosDoGrupo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SkuId,

        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    Write-LogMessage -Message "Iniciando gestão dos usuários do grupo..." -Type INFO
    try {
        # Inicializar contadores
        $addedCount = 0
        $removedCount = 0

        # Obter o grupo
        $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        Write-LogMessage -Message "Grupo encontrado: $($group.DisplayName)" -Type INFO

        # Obter todos os usuários licenciados
        Write-LogMessage -Message "Obtendo usuários com a licença selecionada..." -Type INFO
        $usersWithLicense = @(Get-MgUser -All -Filter "assignedLicenses/any(x:x/skuId eq $SkuId)" -Select "Id,DisplayName,UserPrincipalName")

        # Obter membros atuais do grupo
        Write-LogMessage -Message "Obtendo membros atuais do grupo..." -Type INFO
        $currentMembers = @(Get-MgGroupMember -GroupId $GroupId -All | Select-Object Id, AdditionalProperties)

        # Identificar membros para adicionar e remover
        $usersWithLicenseIds = @($usersWithLicense | Select-Object -ExpandProperty Id)
        $currentMemberIds = @($currentMembers | Select-Object -ExpandProperty Id)

        # Usuários para adicionar (têm licença mas não estão no grupo)
        $usersToAdd = @($usersWithLicense | Where-Object { $_.Id -notin $currentMemberIds })

        # Usuários para remover (estão no grupo mas não têm licença)
        $usersToRemove = @($currentMembers | Where-Object { $_.Id -notin $usersWithLicenseIds })

        # Adicionar usuários ao grupo
        foreach ($user in $usersToAdd) {
            try {
                New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id
                Write-LogMessage -Message "Usuário '$($user.DisplayName)' ($($user.UserPrincipalName)) adicionado ao grupo." -Type SUCCESS
                $addedCount++
            } catch {
                Write-LogMessage -Message "Erro ao adicionar o usuário '$($user.DisplayName)' ao grupo: $_" -Type ERROR
            }
        }

        # Remover usuários do grupo
        foreach ($user in $usersToRemove) {
            try {
                Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id
                $userName = (Get-MgUser -UserId $user.Id).DisplayName
                Write-LogMessage -Message "Usuário '$userName' removido do grupo." -Type SUCCESS
                $removedCount++
            } catch {
                Write-LogMessage -Message "Erro ao remover o usuário com ID '$($user.Id)' do grupo: $_" -Type ERROR
            }
        }

        Write-LogMessage -Message "Gestão dos usuários do grupo concluída." -Type SUCCESS
        
        # Exibir resumo
        Write-LogMessage -Message "Resumo da operação:" -Type INFO
        Write-LogMessage -Message "Usuários adicionados: $addedCount" -Type INFO
        Write-LogMessage -Message "Usuários removidos: $removedCount" -Type INFO
        Write-LogMessage -Message "Total de membros atual: $((Get-MgGroupMember -GroupId $GroupId -All).Count)" -Type INFO

    } catch {
        Write-LogMessage -Message "Erro ao gerenciar os usuários do grupo: $_" -Type ERROR
    }
}

# Função principal para execução do script
function Main {
    try {
        # Importar módulos necessários
        Import-RequiredModules
        
        # Conectar ao Graph
        Connect-GraphPersistent

        do {
            $licenseMapping = Listar-Licencas
            if (-not $licenseMapping) {
                return
            }

            # Selecionar a licença
            do {
                $licenseChoice = Read-Host "Digite o número da licença que deseja processar"
                $isValidLicense = [int]::TryParse($licenseChoice, [ref]$null) -and $licenseMapping.ContainsKey([int]$licenseChoice)
                if (-not $isValidLicense) {
                    Write-LogMessage -Message "Escolha inválida para licença. Tente novamente." -Type ERROR
                }
            } while (-not $isValidLicense)
            
            $selectedLicense = $licenseMapping[[int]$licenseChoice]
            Write-LogMessage -Message "Licença selecionada: $($selectedLicense.SkuPartNumber) | Sku ID: $($selectedLicense.SkuId)" -Type INFO

            # Solicitar o ID do grupo
            $groupId = Read-Host "Digite o ID do grupo que deseja atualizar"

            # Gerenciar a associação do grupo
            Gerenciar-UsuariosDoGrupo -SkuId $selectedLicense.SkuId -GroupId $groupId

            # Perguntar ao usuário se deseja continuar ou sair
            $continuar = Read-Host "Deseja continuar gerenciando outro grupo ou licença? (Sim/Não)"
        } while ($continuar -match '^(S|s)(im)?$')
    } catch {
        Write-LogMessage -Message "Erro crítico na execução do script: $_" -Type ERROR
    } finally {
        # Desconectar do Microsoft Graph se uma sessão for iniciada
        if (Get-MgContext) {
            Disconnect-MgGraph -Confirm:$false
        }
        Write-LogMessage -Message "Execução do script finalizada" -Type INFO
    }
}

# Execução da função principal
Main