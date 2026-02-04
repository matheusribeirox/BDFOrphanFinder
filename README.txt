============================================
  BDFOrphanFinder - Identificador de Arquivos BDF Órfãos
============================================

REQUISITOS:
-----------
- .NET 6.0 SDK ou superior (para compilar)
- Windows 10/11

COMO COMPILAR:
--------------
1. Instale o .NET 6.0 SDK: https://dotnet.microsoft.com/download
2. Execute o arquivo "build.bat"
3. O executável sera gerado em: .\publish\BDFOrphanFinder.exe

COMO USAR:
----------
1. Execute BDFOrphanFinder.exe
2. Preencha os campos:
   - Diretório BDOC: Caminho base do BDOC (ex: C:\Benner\BDOC)
   - Sistema: Nome do sistema a processar (ex: ERP_PROD)
   - Servidor SQL: Endereço do servidor SQL Server
   - Banco de Dados: Nome do banco de dados
   - Usuário: Usuário do SQL Server
   - Senha: Senha do usuário
   - Método de Busca: Escolha entre .NET ou Robocopy
3. Clique em "Iniciar" para começar o processamento
4. Acompanhe o progresso na aba "Log"
5. Veja os arquivos órfãos na aba "Arquivos Órfãos"
6. Clique em "Exportar" para salvar o relatório

RECURSOS:
---------
- Interface gráfica intuitiva
- Barra de progresso em tempo real
- Log detalhado do processamento
- Exportação de relatório em TXT ou CSV
- Salva configurações automaticamente
- Suporte a cancelamento
- Grid com lista de arquivos órfãos

VANTAGENS SOBRE O SCRIPT POWERSHELL:
------------------------------------
- Executável único
- Interface visual mais amigável
- Melhor controle de progresso
- Código mais performático (C# nativo)
- Fácil distribuição

============================================
