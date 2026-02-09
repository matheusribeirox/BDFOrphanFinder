# BDFOrphanFinder

Ferramenta Windows para identificar arquivos BDF órfãos (sem registro correspondente no banco de dados) em estruturas de diretório BDOC do sistema Benner.

## Bancos de dados suportados

- **SQL Server**
- **Oracle**

## Problema

Com o tempo, arquivos `.BDF` podem se acumular nos diretórios BDOC sem possuir um registro `HANDLE` correspondente na base de dados. Esses arquivos órfãos consomem espaço em disco desnecessariamente. Esta ferramenta automatiza a identificação desses arquivos.

## Como funciona

O processo ocorre em fases:

1. **Configuração automática**: O usuário informa o servidor de aplicação Benner. A ferramenta conecta via TCP (portas 5331 e 5337) para obter automaticamente as connection strings, descobrir os diretórios BDOC e listar os sistemas disponíveis.
2. **Fase 1 - Enumeração de arquivos**: Percorre os diretórios BDOC buscando arquivos `*.BDF` e extrai o handle numérico do nome de cada arquivo (formato `<nome>_<handle>.BDF`).
3. **Fase 2 - Validação no banco**: Compara os handles extraídos contra a tabela correspondente no banco de dados. Arquivos cujo handle não existe no banco são marcados como órfãos.
4. **Fase 3 - Backup (opcional)**: Se um caminho de backup for informado, copia os arquivos órfãos preservando a estrutura de diretórios.
5. **Fase 4 - Remoção (opcional)**: Se a opção "Remover órfãos" estiver marcada, remove os arquivos do BDOC após o backup bem-sucedido, incluindo as pastas vazias no caminho.

Ao final, um relatório `.txt` é gerado automaticamente na pasta do executável.

## Autoconfiguração via Servidor de Aplicação

A ferramenta se conecta automaticamente ao servidor de aplicação Benner para obter as informações de conexão com o banco de dados:

1. **Porta 5331** - Obtém a connection string do BSERVER
   - Comando: `getssdbadonetconnectionstring`
   - Consulta sistemas: `SELECT NAME FROM SYS_SYSTEMS WHERE LICINFO IS NOT NULL`
   - Consulta diretórios BDOC: `SELECT NAME, DATA FROM SER_SERVICEPARAMS WHERE NAME IN ('SECDIR','PARDIR') AND SERVICE = 4`

2. **Porta 5337** - Obtém a connection string do sistema selecionado
   - Autenticação: `user internal benner`
   - Seleção: `selectsystem <nome_sistema>`
   - Comando: `getadonetconnectionstring`

O tipo de banco (SQL Server ou Oracle) é detectado automaticamente pela connection string retornada.

### Descoberta automática de diretórios BDOC

Ao conectar, a ferramenta consulta a tabela `SER_SERVICEPARAMS` do BSERVER para descobrir automaticamente os diretórios BDOC configurados:

- **PARDIR**: Diretório primário do BDOC
- **SECDIR**: Diretórios secundários (podem ser múltiplos, separados por `;`)

Se os caminhos forem UNC (ex: `\\SERVIDOR\BDOC\...`) e o servidor for a máquina local, a ferramenta resolve automaticamente para o caminho local do compartilhamento via WMI, priorizando acesso local para melhor performance.

### Busca do sistema em dois níveis

A ferramenta procura a pasta do sistema em dois níveis dentro de cada diretório BDOC:

- **Nível 1**: `BDOC\SISTEMA\` (direto na raiz)
- **Nível 2**: `BDOC\subdir\SISTEMA\` (um nível abaixo)

Isso garante compatibilidade com diferentes estruturas de organização do BDOC.

## Estratégias de busca

| Estratégia | Descrição | Requisitos |
|---|---|---|
| **MFT (Master File Table)** | Lê diretamente da MFT do NTFS para enumeração ultra-rápida. Similar ao Everything Search. Ideal para volumes com milhões de arquivos. | Execução como **Administrador** |
| **.NET EnumerateFiles** | Enumeração e validação em paralelo usando `Directory.EnumerateFiles` e `Parallel.ForEachAsync`. Queries em lote com cláusula `IN`. | Nenhum |

## Requisitos

- Windows (x64)
- Acesso de leitura ao diretório BDOC
- Acesso de rede ao servidor de aplicação Benner (portas 5331 e 5337)
- **Para MFT**: Execução como Administrador

### Para build a partir do código-fonte

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download/dotnet/6.0)

### Para executar o binário self-contained

- Nenhuma dependência adicional (runtime incluso)

## Build

```bash
# Build padrão (requer .NET 6 Runtime no destino)
dotnet build -c Release

# Publish self-contained (não requer .NET instalado no destino)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```

O executável self-contained será gerado em:

```
publish/BDFOrphanFinder.exe
```

## Uso

1. Executar `BDFOrphanFinder.exe` (como Administrador se usar MFT)
2. Preencher os campos:
   - **Diretório BDOC**: Preenchido automaticamente ao conectar (pode ser alterado manualmente)
   - **Serv. Aplicação**: IP ou nome do servidor de aplicação Benner
3. Clicar em **Conectar** para:
   - Obter as connection strings automaticamente
   - Descobrir os diretórios BDOC configurados no BSERVER
   - Carregar os sistemas disponíveis
4. Selecionar o **Sistema** no ComboBox
5. Configurar:
   - **Método de Busca**: MFT ou .NET EnumerateFiles
   - **Threads**: Grau de paralelismo para validação no banco
   - **Caminho Backup BDF** (opcional): Diretório para copiar os órfãos encontrados
   - **Remover órfãos** (opcional): Se marcado, remove os arquivos do BDOC
6. Clicar em **Iniciar**

As configurações são salvas automaticamente em `settings.txt` na pasta do executável.

## Remoção de órfãos

A opção **Remover órfãos** permite limpar automaticamente os arquivos órfãos do BDOC:

### Com backup configurado

1. Marque a opção "Remover órfãos" na interface
2. Configure um caminho de backup
3. Após a Fase 2, uma mensagem de confirmação será exibida com:
   - Quantidade de arquivos órfãos encontrados
   - Tamanho total ocupado
   - Caminho de destino do backup
4. Somente arquivos copiados com sucesso serão removidos do BDOC

### Sem backup configurado

1. Marque a opção "Remover órfãos" sem preencher o caminho de backup
2. Uma mensagem de aviso será exibida informando que os arquivos serão removidos permanentemente sem backup
3. Se confirmado, os arquivos serão removidos diretamente

### Comportamento da remoção

- Pastas vazias no caminho do arquivo (ex: `4F$\F8$\`) também são removidas
- A pasta da tabela e diretórios superiores são preservados
- Arquivos que falharam no backup permanecem intactos
- Um log de auditoria é gerado automaticamente (`auditoria_remocao_YYYYMMDD_HHmmss.txt`)

**ATENÇÃO**: A remoção sem backup é irreversível. Certifique-se de que deseja prosseguir antes de confirmar.

## Saída

- **Log em tela**: Acompanhamento em tempo real com cores por severidade
- **Relatório .txt**: Gerado automaticamente ao final em `arquivos_orfaos_YYYYMMDD_HHmmss.txt` contendo:
  - Log completo da execução
  - Lista de todos os arquivos órfãos com caminho, tabela, handle e tamanho
  - Métricas de performance (duração por fase, throughput)
  - Estatísticas de backup e remoção (se aplicável)

## Diferenças SQL Server vs Oracle

| Aspecto | SQL Server | Oracle |
|---|---|---|
| Escaping de tabelas | `[TABELA] WITH (NOLOCK)` | `"TABELA"` (uppercase) |
| Limite cláusula IN | 2000 valores | 1000 valores |
| Porta padrão | 1433 | 1521 |

## Estrutura do projeto

```
BDFOrphanFinder/
  BDFOrphanFinder.sln        # Solution
  BDFOrphanFinder.csproj     # Projeto .NET 6 WinForms
  Program.cs                 # Entry point
  MainForm.cs                # UI e lógica de processamento
```

## Dependências

- `System.Data.SqlClient` 4.8.6 (conexão SQL Server)
- `Oracle.ManagedDataAccess.Core` 23.x (conexão Oracle)
- `System.Management` 8.0.0 (resolução de caminhos UNC para locais via WMI)

## Autor

Benner
