============================================
  BDFOrphanFinder - Identificador de Arquivos BDF Orfaos
============================================

BANCOS DE DADOS SUPORTADOS:
---------------------------
- SQL Server
- Oracle

REQUISITOS:
-----------
- Windows 10/11 (x64)
- .NET 6.0 SDK ou superior (para compilar)
- Execucao como Administrador (para metodo MFT)

COMO COMPILAR:
--------------
1. Instale o .NET 6.0 SDK: https://dotnet.microsoft.com/download
2. Execute: dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
3. O executavel sera gerado em: .\publish\BDFOrphanFinder.exe

COMO USAR:
----------
1. Execute BDFOrphanFinder.exe (como Administrador se usar MFT)
2. Preencha os campos:
   - Diretorio BDOC: Caminho base do BDOC (ex: B:\bdoc)
   - Sistema: Nome do sistema a processar (ex: SISCON)
   - Tipo Banco: SQL Server ou Oracle
   - Servidor: Endereco do servidor de banco de dados
   - Banco de Dados / Service Name: Nome do banco (SQL) ou Service Name (Oracle)
   - Porta: Porta do Oracle (apenas Oracle, padrao 1521)
   - Usuario: Usuario do banco de dados
   - Senha: Senha do usuario
   - Metodo de Busca: MFT (requer Admin) ou .NET EnumerateFiles
   - Threads: Grau de paralelismo
   - Caminho Backup BDF: (opcional) Diretorio para copiar orfaos
   - Remover orfaos: (opcional) Remove arquivos do BDOC apos backup
3. Clique em "Iniciar" para comecar o processamento
4. Se "Remover orfaos" estiver marcado, confirme a operacao quando solicitado
5. Acompanhe o progresso no log em tempo real
6. O relatorio e gerado automaticamente ao final

METODOS DE BUSCA:
-----------------
- MFT (Master File Table): Le diretamente da MFT do NTFS. Ultra-rapido
  para milhoes de arquivos. REQUER execucao como Administrador.

- .NET EnumerateFiles: Usa Directory.EnumerateFiles com processamento
  paralelo. Nao requer privilegios especiais.

RECURSOS:
---------
- Suporte a SQL Server e Oracle
- Interface grafica intuitiva
- Barra de progresso em tempo real
- Log detalhado do processamento
- Exportacao automatica de relatorio TXT
- Salva configuracoes automaticamente
- Suporte a cancelamento
- Backup opcional dos arquivos orfaos
- Remocao opcional dos orfaos apos backup (com confirmacao)

REMOCAO DE ORFAOS:
------------------
A opcao "Remover orfaos" permite limpar os arquivos do BDOC apos backup:
- Requer caminho de backup configurado
- Exibe confirmacao antes de prosseguir
- Remove apenas arquivos copiados com sucesso
- Remove pastas vazias no caminho (ex: 4F$\F8$\)
- Preserva a pasta da tabela e diretorios superiores
- Arquivos com falha no backup permanecem intactos
ATENCAO: Operacao irreversivel!

DIFERENCAS SQL SERVER vs ORACLE:
--------------------------------
- SQL Server: Usa [TABELA] WITH (NOLOCK), batch de 2000
- Oracle: Usa "TABELA" (uppercase), batch de 1000

============================================
