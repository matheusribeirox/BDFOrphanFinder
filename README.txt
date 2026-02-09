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
- Acesso de rede ao servidor de aplicacao Benner (portas 5331 e 5337)
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
   - Serv. Aplicacao: IP ou nome do servidor de aplicacao Benner
3. Clique em "Conectar" para carregar os sistemas disponiveis
4. Selecione o Sistema no ComboBox
5. Configure:
   - Metodo de Busca: MFT (requer Admin) ou .NET EnumerateFiles
   - Threads: Grau de paralelismo
   - Caminho Backup BDF: (opcional) Diretorio para copiar orfaos
   - Remover orfaos: (opcional) Remove arquivos do BDOC apos backup
6. Clique em "Iniciar" para comecar o processamento
7. Se "Remover orfaos" estiver marcado, confirme a operacao quando solicitado
8. Acompanhe o progresso no log em tempo real
9. O relatorio e gerado automaticamente ao final

AUTO-CONFIGURACAO:
------------------
A ferramenta conecta automaticamente ao servidor de aplicacao Benner:
- Porta 5331: Obtem connection string do BSERVER para listar sistemas
- Porta 5337: Obtem connection string do sistema selecionado
- Tipo de banco (SQL Server/Oracle) detectado automaticamente
- Nao e necessario informar servidor de banco, usuario ou senha

METODOS DE BUSCA:
-----------------
- MFT (Master File Table): Le diretamente da MFT do NTFS. Ultra-rapido
  para milhoes de arquivos. REQUER execucao como Administrador.

- .NET EnumerateFiles: Usa Directory.EnumerateFiles com processamento
  paralelo. Nao requer privilegios especiais.

RECURSOS:
---------
- Auto-configuracao via servidor de aplicacao Benner
- Suporte a SQL Server e Oracle (deteccao automatica)
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
