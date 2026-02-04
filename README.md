# BDFOrphanFinder

Ferramenta Windows para identificar arquivos BDF orfaos (sem registro correspondente no banco de dados SQL Server) em estruturas de diretorio BDOC do sistema Benner.

## Problema

Com o tempo, arquivos `.BDF` podem se acumular nos diretorios BDOC sem possuir um registro `HANDLE` correspondente na base de dados. Esses arquivos orfaos consomem espaco em disco desnecessariamente. Esta ferramenta automatiza a identificacao desses arquivos.

## Como funciona

O processo ocorre em fases:

1. **Fase 1 - Enumeracao de arquivos**: Percorre os diretorios BDOC buscando arquivos `*.BDF` e extrai o handle numerico do nome de cada arquivo (formato `<nome>_<handle>.BDF`).
2. **Fase 2 - Validacao no banco**: Compara os handles extraidos contra a tabela correspondente no SQL Server. Arquivos cujo handle nao existe no banco sao marcados como orfaos.
3. **Fase 3 - Backup (opcional)**: Se um caminho de backup for informado, copia os arquivos orfaos preservando a estrutura de diretorios.

Ao final, um relatorio `.txt` e gerado automaticamente na pasta do executavel.

## Estrategias de busca

| Estrategia | Descricao |
|---|---|
| **Hibrido Paralelo** (Recomendado) | Enumeracao e validacao em paralelo usando `Parallel.ForEachAsync`. Queries em lote com clausula `IN` (batches de 2000). Mais rapido e eficiente. |
| **.NET EnumerateFiles** | Processamento sequencial usando `Directory.EnumerateFiles`. Uma query por tabela trazendo todos os handles. Compatibilidade com ambientes restritos. |
| **Robocopy** | Usa o utilitario nativo `robocopy.exe /L` para listar arquivos. Validacao sequencial no banco. Util para diretorios muito grandes. |

## Requisitos

- Windows (x64)
- Acesso de leitura ao diretorio BDOC
- Acesso ao SQL Server (usuario/senha com permissao de SELECT)

### Para build a partir do codigo-fonte

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download/dotnet/6.0)

### Para executar o binario self-contained

- Nenhuma dependencia adicional (runtime incluso)

## Build

```bash
# Build padrao (requer .NET 6 Runtime no destino)
dotnet build -c Release

# Publish self-contained (nao requer .NET instalado no destino)
dotnet publish -c Release -r win-x64 --self-contained true
```

O executavel self-contained sera gerado em:

```
bin/Release/net6.0-windows/win-x64/publish/
```

## Uso

1. Executar `BDFOrphanFinder.exe`
2. Preencher os campos:
   - **Diretorio BDOC**: Caminho raiz do BDOC (ex: `\\servidor\BDOC`)
   - **Sistema**: Nome do sistema (ex: `SENIOR`, `WES`)
   - **Servidor SQL**: Endereco do SQL Server
   - **Banco de Dados**: Nome do banco
   - **Usuario / Senha**: Credenciais SQL Server
   - **Metodo de Busca**: Escolher a estrategia de enumeracao
   - **Threads**: Grau de paralelismo (apenas para Hibrido Paralelo)
   - **Caminho Backup BDF** (opcional): Diretorio para copiar os orfaos encontrados
3. Clicar em **Iniciar**

As configuracoes sao salvas automaticamente em `settings.txt` na pasta do executavel.

## Saida

- **Log em tela**: Acompanhamento em tempo real com cores por severidade
- **Relatorio .txt**: Gerado automaticamente ao final em `arquivos_orfaos_YYYYMMDD_HHmmss.txt` contendo:
  - Log completo da execucao
  - Lista de todos os arquivos orfaos com caminho, tabela, handle e tamanho
  - Metricas de performance (duracao por fase, throughput)

## Estrutura do projeto

```
BDFOrphanFinder/
  BDFOrphanFinder.sln        # Solution
  BDFOrphanFinder.csproj      # Projeto .NET 6 WinForms
  Program.cs                  # Entry point
  MainForm.cs                 # UI e logica de processamento
```

## Dependencias

- `System.Data.SqlClient` 4.8.6 (conexao SQL Server)

## Autor

Benner
