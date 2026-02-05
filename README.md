# BDFOrphanFinder

Ferramenta Windows para identificar arquivos BDF orfaos (sem registro correspondente no banco de dados) em estruturas de diretorio BDOC do sistema Benner.

## Bancos de dados suportados

- **SQL Server**
- **Oracle**

## Problema

Com o tempo, arquivos `.BDF` podem se acumular nos diretorios BDOC sem possuir um registro `HANDLE` correspondente na base de dados. Esses arquivos orfaos consomem espaco em disco desnecessariamente. Esta ferramenta automatiza a identificacao desses arquivos.

## Como funciona

O processo ocorre em fases:

1. **Fase 1 - Enumeracao de arquivos**: Percorre os diretorios BDOC buscando arquivos `*.BDF` e extrai o handle numerico do nome de cada arquivo (formato `<nome>_<handle>.BDF`).
2. **Fase 2 - Validacao no banco**: Compara os handles extraidos contra a tabela correspondente no banco de dados. Arquivos cujo handle nao existe no banco sao marcados como orfaos.
3. **Fase 3 - Backup (opcional)**: Se um caminho de backup for informado, copia os arquivos orfaos preservando a estrutura de diretorios.
4. **Fase 4 - Remocao (opcional)**: Se a opcao "Remover orfaos" estiver marcada, remove os arquivos do BDOC apos o backup bem-sucedido, incluindo as pastas vazias no caminho.

Ao final, um relatorio `.txt` e gerado automaticamente na pasta do executavel.

## Estrategias de busca

| Estrategia | Descricao | Requisitos |
|---|---|---|
| **MFT (Master File Table)** | Le diretamente da MFT do NTFS para enumeracao ultra-rapida. Similar ao Everything Search. Ideal para volumes com milhoes de arquivos. | Execucao como **Administrador** |
| **.NET EnumerateFiles** | Enumeracao e validacao em paralelo usando `Directory.EnumerateFiles` e `Parallel.ForEachAsync`. Queries em lote com clausula `IN`. | Nenhum |

## Requisitos

- Windows (x64)
- Acesso de leitura ao diretorio BDOC
- Acesso ao banco de dados (SQL Server ou Oracle) com permissao de SELECT
- **Para MFT**: Execucao como Administrador

### Para build a partir do codigo-fonte

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download/dotnet/6.0)

### Para executar o binario self-contained

- Nenhuma dependencia adicional (runtime incluso)

## Build

```bash
# Build padrao (requer .NET 6 Runtime no destino)
dotnet build -c Release

# Publish self-contained (nao requer .NET instalado no destino)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```

O executavel self-contained sera gerado em:

```
publish/BDFOrphanFinder.exe
```

## Uso

1. Executar `BDFOrphanFinder.exe` (como Administrador se usar MFT)
2. Preencher os campos:
   - **Diretorio BDOC**: Caminho raiz do BDOC (ex: `B:\bdoc`)
   - **Sistema**: Nome do sistema (ex: `SISCON`, `WES`)
   - **Tipo Banco**: SQL Server ou Oracle
   - **Servidor**: Endereco do servidor de banco de dados
   - **Banco de Dados / Service Name**: Nome do banco (SQL Server) ou Service Name (Oracle)
   - **Porta**: Porta do Oracle (apenas para Oracle, padrao 1521)
   - **Usuario / Senha**: Credenciais do banco
   - **Metodo de Busca**: MFT ou .NET EnumerateFiles
   - **Threads**: Grau de paralelismo para validacao no banco
   - **Caminho Backup BDF** (opcional): Diretorio para copiar os orfaos encontrados
   - **Remover orfaos** (opcional): Se marcado, remove os arquivos do BDOC apos o backup
3. Clicar em **Iniciar**
4. Se "Remover orfaos" estiver marcado, uma confirmacao sera solicitada antes de prosseguir

As configuracoes sao salvas automaticamente em `settings.txt` na pasta do executavel.

## Formato de conexao Oracle

A conexao Oracle usa o formato TNS:

```
Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=servidor)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=nome_servico)));User Id=usuario;Password=senha;
```

## Remocao de orfaos

A opcao **Remover orfaos** permite limpar automaticamente os arquivos orfaos do BDOC apos o backup:

1. Marque a opcao "Remover orfaos" na interface
2. Configure um caminho de backup valido (obrigatorio quando remocao esta ativa)
3. Apos a Fase 2, uma mensagem de confirmacao sera exibida com:
   - Quantidade de arquivos orfaos encontrados
   - Tamanho total ocupado
   - Caminho de destino do backup
4. Somente arquivos copiados com sucesso serao removidos do BDOC
5. Pastas vazias no caminho do arquivo (ex: `4F$\F8$\`) tambem sao removidas
6. A pasta da tabela e diretorios superiores sao preservados
7. Arquivos que falharam no backup permanecem intactos

**ATENCAO**: Esta operacao e irreversivel. Certifique-se de que o backup foi realizado corretamente antes de confirmar.

## Saida

- **Log em tela**: Acompanhamento em tempo real com cores por severidade
- **Relatorio .txt**: Gerado automaticamente ao final em `arquivos_orfaos_YYYYMMDD_HHmmss.txt` contendo:
  - Log completo da execucao
  - Lista de todos os arquivos orfaos com caminho, tabela, handle e tamanho
  - Metricas de performance (duracao por fase, throughput)
  - Estatisticas de backup e remocao (se aplicavel)

## Diferencas SQL Server vs Oracle

| Aspecto | SQL Server | Oracle |
|---|---|---|
| Escaping de tabelas | `[TABELA] WITH (NOLOCK)` | `"TABELA"` (uppercase) |
| Limite clausula IN | 2000 valores | 1000 valores |
| Porta padrao | 1433 | 1521 |

## Estrutura do projeto

```
BDFOrphanFinder/
  BDFOrphanFinder.sln        # Solution
  BDFOrphanFinder.csproj     # Projeto .NET 6 WinForms
  Program.cs                 # Entry point
  MainForm.cs                # UI e logica de processamento
```

## Dependencias

- `System.Data.SqlClient` 4.8.6 (conexao SQL Server)
- `Oracle.ManagedDataAccess.Core` 23.x (conexao Oracle)

## Autor

Benner
