using System.Collections.Concurrent;
using System.Data.SqlClient;
using System.Text;

namespace BDFOrphanFinder;

public partial class MainForm : Form
{
    private TextBox txtBdocPath = null!;
    private TextBox txtSystemName = null!;
    private TextBox txtServer = null!;
    private TextBox txtDatabase = null!;
    private TextBox txtUsername = null!;
    private TextBox txtPassword = null!;
    private ComboBox cmbSearchMethod = null!;
    private NumericUpDown numParallelism = null!;
    private TextBox txtBackupPath = null!;
    private Button btnBrowse = null!;
    private Button btnBrowseBackup = null!;
    private Button btnStart = null!;
    private Button btnCancel = null!;
    private ProgressBar progressBar = null!;
    private Label lblStatus = null!;
    private Label lblProgress = null!;
    private RichTextBox txtLog = null!;

    private CancellationTokenSource? _cts;
    private ConcurrentBag<OrphanFile> _orphanFiles = new();
    private int _totalFilesAnalyzed = 0;
    private int _filesWithRecord = 0;
    private long _totalOrphanSize = 0;

    // Métricas de performance
    private TimeSpan _phase1Duration;
    private TimeSpan _phase2Duration;
    private TimeSpan _phase3Duration;
    private int _totalTablesProcessed = 0;
    private int _totalFilesCopied = 0;
    private int _totalCopyErrors = 0;

    public MainForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Identificador de Arquivos BDF Orfaos";
        this.Size = new Size(900, 700);
        this.MinimumSize = new Size(800, 600);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Font = new Font("Segoe UI", 9.5f);
        this.BackColor = Color.FromArgb(248, 249, 252);

        var mainPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            RowCount = 3,
            ColumnCount = 1,
            BackColor = Color.FromArgb(248, 249, 252)
        };
        mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // ===== CONFIG PANEL =====
        var configGroup = new GroupBox
        {
            Text = "Configuracoes",
            Dock = DockStyle.Fill,
            AutoSize = true,
            Padding = new Padding(15, 20, 15, 10),
            ForeColor = Color.FromArgb(0, 120, 212),
            Font = new Font("Segoe UI Semibold", 10f),
            BackColor = Color.White
        };

        var configPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 6,
            RowCount = 7,
            BackColor = Color.White
        };
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));  // Col 0: Label
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); // Col 1: TextBox
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));  // Col 2: Button ...
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));  // Col 3: Label
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // Col 4: TextBox
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));  // Col 5: Spacer

        // Estilo comum para labels
        var labelColor = Color.FromArgb(51, 51, 51);
        var labelFont = new Font("Segoe UI", 9f);
        var labelMarginLeft = new Padding(3, 8, 3, 0);
        var labelMarginRight = new Padding(10, 8, 3, 0);
        var textBoxFont = new Font("Segoe UI", 9.5f);

        // Row 0: BDOC Path + Sistema
        configPanel.Controls.Add(new Label { Text = "Diretorio BDOC:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 0);
        txtBdocPath = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtBdocPath, 1, 0);
        btnBrowse = new Button
        {
            Text = "...", Width = 32, Height = 26,
            FlatStyle = FlatStyle.Flat, BackColor = Color.White,
            ForeColor = labelColor, Font = labelFont, Cursor = Cursors.Hand
        };
        btnBrowse.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
        btnBrowse.FlatAppearance.BorderSize = 1;
        btnBrowse.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 240);
        btnBrowse.Click += BtnBrowse_Click;
        configPanel.Controls.Add(btnBrowse, 2, 0);
        configPanel.Controls.Add(new Label { Text = "Sistema:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 0);
        txtSystemName = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtSystemName, 4, 0);

        // Row 1: Server + Database
        configPanel.Controls.Add(new Label { Text = "Servidor SQL:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 1);
        txtServer = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtServer, 1, 1);
        configPanel.Controls.Add(new Label { Text = "Banco de Dados:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 1);
        txtDatabase = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtDatabase, 4, 1);

        // Row 2: Username + Password
        configPanel.Controls.Add(new Label { Text = "Usuario:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 2);
        txtUsername = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtUsername, 1, 2);
        configPanel.Controls.Add(new Label { Text = "Senha:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 2);
        txtPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtPassword, 4, 2);

        // Row 3: Search Method + Parallelism
        configPanel.Controls.Add(new Label { Text = "Metodo de Busca:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 3);
        cmbSearchMethod = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            Font = textBoxFont
        };
        cmbSearchMethod.Items.AddRange(new object[] { "Hibrido Paralelo (Recomendado)", ".NET EnumerateFiles", "Robocopy" });
        cmbSearchMethod.SelectedIndex = 0;
        configPanel.Controls.Add(cmbSearchMethod, 1, 3);

        configPanel.Controls.Add(new Label { Text = "Threads:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 3);
        numParallelism = new NumericUpDown
        {
            Minimum = 1,
            Maximum = Environment.ProcessorCount * 2,
            Value = Math.Min(4, Environment.ProcessorCount),
            Width = 60,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            Font = textBoxFont
        };
        configPanel.Controls.Add(numParallelism, 4, 3);

        // Row 4: Backup Path
        configPanel.Controls.Add(new Label { Text = "Caminho Backup BDF:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 4);
        txtBackupPath = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtBackupPath, 1, 4);
        btnBrowseBackup = new Button
        {
            Text = "...", Width = 32, Height = 26,
            FlatStyle = FlatStyle.Flat, BackColor = Color.White,
            ForeColor = labelColor, Font = labelFont, Cursor = Cursors.Hand
        };
        btnBrowseBackup.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
        btnBrowseBackup.FlatAppearance.BorderSize = 1;
        btnBrowseBackup.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 240);
        btnBrowseBackup.Click += BtnBrowseBackup_Click;
        configPanel.Controls.Add(btnBrowseBackup, 2, 4);

        // Row 5: Buttons
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true
        };

        btnStart = new Button
        {
            Text = "Iniciar", Width = 100, Height = 32,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 9.5f),
            Cursor = Cursors.Hand
        };
        btnStart.FlatAppearance.BorderSize = 0;
        btnStart.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 180);
        btnStart.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 84, 153);
        btnStart.Click += BtnStart_Click;

        btnCancel = new Button
        {
            Text = "Cancelar", Width = 100, Height = 32,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(240, 240, 240),
            ForeColor = Color.FromArgb(51, 51, 51),
            Font = new Font("Segoe UI", 9.5f),
            Cursor = Cursors.Hand,
            Enabled = false
        };
        btnCancel.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnCancel.FlatAppearance.BorderSize = 1;
        btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 225);
        btnCancel.FlatAppearance.MouseDownBackColor = Color.FromArgb(210, 210, 210);
        btnCancel.Click += BtnCancel_Click;

        buttonPanel.Controls.Add(btnStart);
        buttonPanel.Controls.Add(btnCancel);

        configPanel.Controls.Add(buttonPanel, 1, 5);
        configPanel.SetColumnSpan(buttonPanel, 4);

        configGroup.Controls.Add(configPanel);
        mainPanel.Controls.Add(configGroup, 0, 0);

        // ===== LOG PANEL =====
        txtLog = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            Font = new Font("Consolas", 9),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.White,
            Margin = new Padding(0, 8, 0, 8),
            BorderStyle = BorderStyle.None
        };

        mainPanel.Controls.Add(txtLog, 0, 1);

        // ===== STATUS PANEL =====
        var statusPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            AutoSize = true,
            BackColor = Color.FromArgb(248, 249, 252)
        };
        statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        progressBar = new ProgressBar { Dock = DockStyle.Fill, Height = 22 };
        statusPanel.Controls.Add(progressBar, 0, 0);

        lblProgress = new Label
        {
            Text = "0%", AutoSize = true, Anchor = AnchorStyles.Left,
            ForeColor = Color.FromArgb(0, 120, 212),
            Font = new Font("Segoe UI Semibold", 9.5f)
        };
        statusPanel.Controls.Add(lblProgress, 1, 0);

        lblStatus = new Label
        {
            Text = "Aguardando...", Dock = DockStyle.Fill, AutoSize = true,
            ForeColor = Color.FromArgb(100, 100, 100),
            Font = new Font("Segoe UI", 9f)
        };
        statusPanel.Controls.Add(lblStatus, 0, 1);
        statusPanel.SetColumnSpan(lblStatus, 2);

        mainPanel.Controls.Add(statusPanel, 0, 2);

        this.Controls.Add(mainPanel);
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Selecione o diretorio base do BDOC",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtBdocPath.Text = dialog.SelectedPath;
        }
    }

    private void BtnBrowseBackup_Click(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Selecione o diretorio de backup dos BDF orfaos",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtBackupPath.Text = dialog.SelectedPath;
        }
    }

    private async void BtnStart_Click(object? sender, EventArgs e)
    {
        if (!ValidateInputs()) return;

        SaveSettings();

        _cts = new CancellationTokenSource();
        _orphanFiles = new ConcurrentBag<OrphanFile>();
        _totalFilesAnalyzed = 0;
        _filesWithRecord = 0;
        _totalOrphanSize = 0;
        _totalTablesProcessed = 0;
        _phase1Duration = TimeSpan.Zero;
        _phase2Duration = TimeSpan.Zero;
        _phase3Duration = TimeSpan.Zero;
        _totalFilesCopied = 0;
        _totalCopyErrors = 0;
        txtLog.Clear();

        SetUIState(running: true);

        try
        {
            await ProcessOrphanFilesAsync(_cts.Token);
        }
        catch (OperationCanceledException)
        {
            Log("Operacao cancelada pelo usuario.", Color.Yellow);
        }
        catch (Exception ex)
        {
            Log($"ERRO: {ex.Message}", Color.Red);
            MessageBox.Show($"Erro durante o processamento:\n{ex.Message}", "Erro",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetUIState(running: false);
            _cts?.Dispose();
            _cts = null;
        }
    }

    private void BtnCancel_Click(object? sender, EventArgs e)
    {
        _cts?.Cancel();
        lblStatus.Text = "Cancelando...";
    }

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(txtBdocPath.Text) || !Directory.Exists(txtBdocPath.Text))
        {
            MessageBox.Show("Diretorio BDOC invalido ou nao existe.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (string.IsNullOrWhiteSpace(txtSystemName.Text))
        {
            MessageBox.Show("Nome do sistema e obrigatorio.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (string.IsNullOrWhiteSpace(txtServer.Text) ||
            string.IsNullOrWhiteSpace(txtDatabase.Text) ||
            string.IsNullOrWhiteSpace(txtUsername.Text))
        {
            MessageBox.Show("Dados de conexao SQL Server incompletos.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (!string.IsNullOrWhiteSpace(txtBackupPath.Text) && !Directory.Exists(txtBackupPath.Text))
        {
            MessageBox.Show("Diretorio de backup invalido ou nao existe.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        return true;
    }

    private void SetUIState(bool running)
    {
        btnStart.Enabled = !running;
        btnCancel.Enabled = running;
        btnBrowse.Enabled = !running;
        btnBrowseBackup.Enabled = !running;
        txtBdocPath.Enabled = !running;
        txtBackupPath.Enabled = !running;
        txtSystemName.Enabled = !running;
        txtServer.Enabled = !running;
        txtDatabase.Enabled = !running;
        txtUsername.Enabled = !running;
        txtPassword.Enabled = !running;
        cmbSearchMethod.Enabled = !running;
        numParallelism.Enabled = !running;

        // Ajustar cor do botao Iniciar conforme estado
        btnStart.BackColor = running
            ? Color.FromArgb(180, 180, 180)
            : Color.FromArgb(0, 120, 212);

        if (!running)
        {
            progressBar.Value = 100;
            lblProgress.Text = "100%";
        }
        else
        {
            progressBar.Value = 0;
            lblProgress.Text = "0%";
        }
    }

    private void Log(string message, Color? color = null)
    {
        if (InvokeRequired)
        {
            Invoke(() => Log(message, color));
            return;
        }

        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.SelectionColor = color ?? Color.White;
        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
        txtLog.ScrollToCaret();
    }

    private void UpdateProgress(int current, int total, string status)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateProgress(current, total, status));
            return;
        }

        int percent = total > 0 ? (int)((double)current / total * 100) : 0;
        progressBar.Value = Math.Min(percent, 100);
        lblProgress.Text = $"{percent}%";
        lblStatus.Text = status;
    }

    private async Task ProcessOrphanFilesAsync(CancellationToken ct)
    {
        var bdocPath = txtBdocPath.Text;
        var systemName = txtSystemName.Text;
        var maxParallelism = (int)numParallelism.Value;
        var searchMethod = cmbSearchMethod.SelectedIndex;

        Log("Iniciando processamento...", Color.Cyan);
        Log($"Diretorio BDOC: {bdocPath}");
        Log($"Sistema: {systemName}");
        Log($"Metodo: {cmbSearchMethod.SelectedItem}");
        Log($"Threads: {maxParallelism}");
        Log("");

        var startTime = DateTime.Now;

        // Find system paths
        Log("Procurando caminhos do sistema...");
        var systemPaths = new List<string>();
        foreach (var subdir in Directory.EnumerateDirectories(bdocPath))
        {
            var systemPath = Path.Combine(subdir, systemName);
            if (Directory.Exists(systemPath))
            {
                systemPaths.Add(systemPath);
                Log($"  [ENCONTRADO] {systemPath}", Color.Green);
            }
        }

        if (systemPaths.Count == 0)
        {
            Log($"ERRO: Sistema '{systemName}' nao encontrado.", Color.Red);
            return;
        }

        // Connection string for parallel connections
        var connectionString = $"Server={txtServer.Text};Database={txtDatabase.Text};User Id={txtUsername.Text};Password={txtPassword.Text};TrustServerCertificate=True;Max Pool Size={maxParallelism + 5};";

        // Test connection
        Log("Testando conexao com banco de dados...");
        using (var testConn = new SqlConnection(connectionString))
        {
            await testConn.OpenAsync(ct);
            Log("Conexao estabelecida com sucesso!", Color.Green);
        }

        // Use hybrid parallel strategy or legacy sequential
        if (searchMethod == 0) // Híbrido Paralelo
        {
            await ProcessWithHybridStrategyAsync(systemPaths, systemName, connectionString, maxParallelism, ct);
        }
        else
        {
            await ProcessWithLegacyStrategyAsync(systemPaths, systemName, connectionString, searchMethod, ct);
        }

        // ================================================================
        // FASE 3: Backup dos arquivos órfãos (se configurado)
        // ================================================================
        var backupPath = txtBackupPath.Text;
        if (!string.IsNullOrWhiteSpace(backupPath) && !_orphanFiles.IsEmpty)
        {
            Log("");
            Log("=== FASE 3: Backup dos arquivos orfaos ===", Color.Cyan);
            var phase3Start = DateTime.Now;

            var orphanList = _orphanFiles.ToList();
            var copied = 0;
            var errors = 0;
            var totalOrphans = orphanList.Count;

            Log($"Copiando {totalOrphans} arquivos para: {backupPath}");

            await Parallel.ForEachAsync(orphanList,
                new ParallelOptions { MaxDegreeOfParallelism = (int)numParallelism.Value, CancellationToken = ct },
                async (orphan, token) =>
                {
                    try
                    {
                        var relativePath = Path.GetRelativePath(bdocPath, orphan.FilePath);
                        var destPath = Path.Combine(backupPath, relativePath);
                        var destDir = Path.GetDirectoryName(destPath)!;

                        Directory.CreateDirectory(destDir);
                        File.Copy(orphan.FilePath, destPath, overwrite: false);

                        var done = Interlocked.Increment(ref copied);
                        if (done % 100 == 0 || done == totalOrphans)
                        {
                            UpdateProgress(done, totalOrphans, $"Fase 3: Backup ({done}/{totalOrphans} arquivos)");
                        }
                    }
                    catch (Exception ex)
                    {
                        Interlocked.Increment(ref errors);
                        Log($"  [ERRO BACKUP] {orphan.FileName}: {ex.Message}", Color.Yellow);
                    }

                    await Task.CompletedTask;
                });

            _phase3Duration = DateTime.Now - phase3Start;
            _totalFilesCopied = copied;
            _totalCopyErrors = errors;

            Log($"Backup concluido: {copied} copiados, {errors} erros", copied > 0 ? Color.Green : Color.Yellow);
            Log($"Fase 3 concluida em: {_phase3Duration:mm\\:ss\\.fff}", Color.Green);
        }

        // Final report
        var elapsed = DateTime.Now - startTime;
        var orphanCount = _orphanFiles.Count;

        Log("", Color.White);
        Log("============================================", Color.Cyan);
        Log("           RELATORIO FINAL", Color.Cyan);
        Log("============================================", Color.Cyan);
        Log($"Total de arquivos analisados: {_totalFilesAnalyzed}");
        Log($"Arquivos COM registro na base: {_filesWithRecord}", Color.Green);
        Log($"Arquivos ORFAOS (sem registro): {orphanCount}", orphanCount > 0 ? Color.Red : Color.Green);
        Log($"Espaco ocupado por orfaos: {FormatSize(_totalOrphanSize)}", Color.Yellow);
        if (_totalFilesCopied > 0 || _totalCopyErrors > 0)
        {
            Log($"Backup: {_totalFilesCopied} copiados, {_totalCopyErrors} erros", Color.Cyan);
        }
        Log("");
        Log("--- METRICAS DE PERFORMANCE ---", Color.Cyan);
        Log($"Fase 1 (Enumeracao de arquivos): {_phase1Duration:hh\\:mm\\:ss\\.fff}");
        Log($"Fase 2 (Validacao no banco): {_phase2Duration:hh\\:mm\\:ss\\.fff}");
        if (_phase3Duration > TimeSpan.Zero)
        {
            Log($"Fase 3 (Backup de orfaos): {_phase3Duration:hh\\:mm\\:ss\\.fff}");
        }
        Log($"Tabelas processadas: {_totalTablesProcessed}");
        Log($"Tempo total: {elapsed:hh\\:mm\\:ss\\.fff}");
        Log($"Throughput: {(_totalFilesAnalyzed / Math.Max(1, elapsed.TotalSeconds)):F0} arquivos/segundo");
        Log("============================================", Color.Cyan);

        // Gerar relatório .txt automaticamente na pasta do executável
        var reportDir = Environment.CurrentDirectory;
        var reportFileName = $"arquivos_orfaos_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        var reportPath = Path.Combine(reportDir, reportFileName);
        ExportResults(reportPath);
    }

    /// <summary>
    /// ESTRATÉGIA HÍBRIDA OTIMIZADA
    /// Fase 1: Enumera todos os arquivos em paralelo, agrupando por tabela
    /// Fase 2: Valida handles no banco em paralelo usando IN clause
    /// </summary>
    private async Task ProcessWithHybridStrategyAsync(
        List<string> systemPaths,
        string systemName,
        string connectionString,
        int maxParallelism,
        CancellationToken ct)
    {
        // ================================================================
        // FASE 1: Enumeração paralela de arquivos por diretório de tabela
        // ================================================================
        Log("");
        Log("=== FASE 1: Enumeracao paralela de arquivos ===", Color.Cyan);
        var phase1Start = DateTime.Now;

        // Descobrir todos os diretórios de tabelas
        var tableDirs = new ConcurrentDictionary<string, ConcurrentBag<string>>();
        foreach (var systemPath in systemPaths)
        {
            foreach (var tableDir in Directory.EnumerateDirectories(systemPath))
            {
                var tableName = Path.GetFileName(tableDir);
                tableDirs.GetOrAdd(tableName, _ => new ConcurrentBag<string>()).Add(tableDir);
            }
        }

        Log($"Tabelas encontradas: {tableDirs.Count}");

        // Estrutura thread-safe para armazenar arquivos por tabela
        var filesByTable = new ConcurrentDictionary<string, ConcurrentBag<FileHandleInfo>>();
        var totalFilesFound = 0;
        var directoriesProcessed = 0;
        var totalDirectories = tableDirs.Values.Sum(b => b.Count);

        // Processar diretórios em paralelo
        var allDirs = tableDirs.SelectMany(kvp => kvp.Value.Select(dir => (Table: kvp.Key, Dir: dir))).ToList();

        await Parallel.ForEachAsync(allDirs,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallelism, CancellationToken = ct },
            async (item, token) =>
            {
                var (tableName, tableDir) = item;

                var bag = filesByTable.GetOrAdd(tableName, _ => new ConcurrentBag<FileHandleInfo>());
                var filesInDir = 0;

                try
                {
                    foreach (var filePath in Directory.EnumerateFiles(tableDir, "*.BDF", SearchOption.AllDirectories))
                    {
                        token.ThrowIfCancellationRequested();

                        if (TryExtractHandle(filePath, out var handle, out var fileName))
                        {
                            bag.Add(new FileHandleInfo(filePath, fileName, handle));
                            filesInDir++;
                        }
                    }
                }
                catch (UnauthorizedAccessException) { }
                catch (DirectoryNotFoundException) { }

                Interlocked.Add(ref totalFilesFound, filesInDir);
                var processed = Interlocked.Increment(ref directoriesProcessed);

                if (processed % 10 == 0 || processed == totalDirectories)
                {
                    UpdateProgress(processed, totalDirectories, $"Fase 1: Enumerando arquivos ({processed}/{totalDirectories} dirs)");
                }

                await Task.CompletedTask; // Para manter async
            });

        _phase1Duration = DateTime.Now - phase1Start;
        _totalFilesAnalyzed = totalFilesFound;

        Log($"Arquivos encontrados: {totalFilesFound} em {tableDirs.Count} tabelas", Color.Green);
        Log($"Fase 1 concluida em: {_phase1Duration:mm\\:ss\\.fff}", Color.Green);

        // ================================================================
        // FASE 2: Validação paralela no banco de dados
        // ================================================================
        Log("");
        Log("=== FASE 2: Validacao paralela no banco de dados ===", Color.Cyan);
        var phase2Start = DateTime.Now;

        var tablesToProcess = filesByTable.Where(kvp => !kvp.Value.IsEmpty).ToList();
        var tablesProcessed = 0;
        var totalTables = tablesToProcess.Count;

        Log($"Tabelas com arquivos para validar: {totalTables}");

        // Processar tabelas em paralelo com conexões do pool
        await Parallel.ForEachAsync(tablesToProcess,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallelism, CancellationToken = ct },
            async (kvp, token) =>
            {
                var tableName = kvp.Key;
                var files = kvp.Value.ToList();
                var handleList = files.Select(f => f.Handle).Distinct().ToList();

                try
                {
                    // Buscar handles existentes no banco usando IN clause em batches
                    var existingHandles = await QueryExistingHandlesAsync(
                        connectionString, tableName, handleList, token);

                    // Identificar órfãos
                    var orphansInTable = 0;
                    foreach (var file in files)
                    {
                        if (!existingHandles.Contains(file.Handle))
                        {
                            long size = 0;
                            try { size = new FileInfo(file.FilePath).Length; } catch { }

                            var orphan = new OrphanFile
                            {
                                Sistema = systemName,
                                Tabela = tableName,
                                Handle = file.Handle,
                                FileName = file.FileName,
                                FilePath = file.FilePath,
                                Size = size
                            };

                            _orphanFiles.Add(orphan);
                            Interlocked.Add(ref _totalOrphanSize, size);
                            orphansInTable++;
                        }
                        else
                        {
                            Interlocked.Increment(ref _filesWithRecord);
                        }
                    }

                    if (orphansInTable > 0)
                    {
                        Log($"  [{tableName}] {files.Count} arquivos, {orphansInTable} orfaos", Color.Red);
                    }
                }
                catch (SqlException ex)
                {
                    Log($"  [AVISO] Erro na tabela {tableName}: {ex.Message}", Color.Yellow);
                }

                var processed = Interlocked.Increment(ref tablesProcessed);
                Interlocked.Increment(ref _totalTablesProcessed);
                UpdateProgress(processed, totalTables, $"Fase 2: Validando ({processed}/{totalTables} tabelas)");
            });

        _phase2Duration = DateTime.Now - phase2Start;
        Log($"Fase 2 concluida em: {_phase2Duration:mm\\:ss\\.fff}", Color.Green);
    }

    /// <summary>
    /// Extrai o handle do nome do arquivo
    /// </summary>
    private static bool TryExtractHandle(string filePath, out int handle, out string fileName)
    {
        handle = 0;
        fileName = Path.GetFileName(filePath);
        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        var lastUnderscore = fileNameWithoutExt.LastIndexOf('_');

        if (lastUnderscore == -1) return false;

        var handleStr = fileNameWithoutExt.AsSpan()[(lastUnderscore + 1)..];
        return int.TryParse(handleStr, out handle);
    }

    /// <summary>
    /// Busca handles existentes no banco usando IN clause em batches
    /// Limite de ~2000 parâmetros por query no SQL Server
    /// </summary>
    private static async Task<HashSet<int>> QueryExistingHandlesAsync(
        string connectionString,
        string tableName,
        List<int> handles,
        CancellationToken ct)
    {
        var existingHandles = new HashSet<int>();
        const int batchSize = 2000; // SQL Server limit

        using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(ct);

        for (int i = 0; i < handles.Count; i += batchSize)
        {
            ct.ThrowIfCancellationRequested();

            var batch = handles.Skip(i).Take(batchSize).ToList();
            var inClause = string.Join(",", batch);

            var query = $"SELECT DISTINCT HANDLE FROM [{tableName}] WITH (NOLOCK) WHERE HANDLE IN ({inClause})";

            using var cmd = new SqlCommand(query, connection);
            cmd.CommandTimeout = 120;

            using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                if (!reader.IsDBNull(0))
                    existingHandles.Add(reader.GetInt32(0));
            }
        }

        return existingHandles;
    }

    /// <summary>
    /// Estratégia legada (sequencial) para compatibilidade
    /// </summary>
    private async Task ProcessWithLegacyStrategyAsync(
        List<string> systemPaths,
        string systemName,
        string connectionString,
        int searchMethod,
        CancellationToken ct)
    {
        Log("");
        Log("=== Processamento Sequencial (Legado) ===", Color.Yellow);
        var phase1Start = DateTime.Now;

        using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(ct);

        // Discover tables
        var tables = new HashSet<string>();
        var tablePaths = new Dictionary<string, List<string>>();

        foreach (var systemPath in systemPaths)
        {
            ct.ThrowIfCancellationRequested();
            foreach (var tableDir in Directory.EnumerateDirectories(systemPath))
            {
                var tableName = Path.GetFileName(tableDir);
                tables.Add(tableName);

                if (!tablePaths.ContainsKey(tableName))
                    tablePaths[tableName] = new List<string>();
                tablePaths[tableName].Add(tableDir);
            }
        }

        Log($"Total de tabelas encontradas: {tables.Count}", Color.Cyan);
        _phase1Duration = DateTime.Now - phase1Start;

        var phase2Start = DateTime.Now;
        int tableIndex = 0;
        int totalTables = tables.Count;

        foreach (var table in tables)
        {
            ct.ThrowIfCancellationRequested();
            tableIndex++;

            UpdateProgress(tableIndex, totalTables, $"Processando tabela {table} ({tableIndex}/{totalTables})");
            Log($"[{tableIndex}/{totalTables}] Processando tabela: {table}", Color.Yellow);

            var handles = new HashSet<int>();
            try
            {
                var query = $"SELECT DISTINCT HANDLE FROM [{table}] WITH (NOLOCK)";
                using var cmd = new SqlCommand(query, connection);
                cmd.CommandTimeout = 300;

                using var reader = await cmd.ExecuteReaderAsync(ct);
                while (await reader.ReadAsync(ct))
                {
                    if (!reader.IsDBNull(0))
                        handles.Add(reader.GetInt32(0));
                }
            }
            catch (SqlException ex)
            {
                Log($"  [AVISO] Erro ao consultar tabela {table}: {ex.Message}", Color.Yellow);
                continue;
            }

            if (handles.Count == 0) continue;

            foreach (var tablePath in tablePaths[table])
            {
                ct.ThrowIfCancellationRequested();
                if (!Directory.Exists(tablePath)) continue;

                var files = searchMethod == 1 ? GetFilesWithDotNet(tablePath) : GetFilesWithRobocopy(tablePath);

                foreach (var filePath in files)
                {
                    ct.ThrowIfCancellationRequested();
                    _totalFilesAnalyzed++;

                    if (!TryExtractHandle(filePath, out var handle, out var fileName)) continue;

                    if (!handles.Contains(handle))
                    {
                        long size = 0;
                        try { size = new FileInfo(filePath).Length; } catch { }

                        var orphan = new OrphanFile
                        {
                            Sistema = systemName,
                            Tabela = table,
                            Handle = handle,
                            FileName = fileName,
                            FilePath = filePath,
                            Size = size
                        };

                        _orphanFiles.Add(orphan);
                        Interlocked.Add(ref _totalOrphanSize, size);
                    }
                    else
                    {
                        _filesWithRecord++;
                    }
                }
            }
            _totalTablesProcessed++;
        }

        _phase2Duration = DateTime.Now - phase2Start;
    }

    private IEnumerable<string> GetBdfFiles(string path)
    {
        if (cmbSearchMethod.SelectedIndex == 1) // Robocopy
        {
            return GetFilesWithRobocopy(path);
        }
        else // .NET EnumerateFiles
        {
            return GetFilesWithDotNet(path);
        }
    }

    private IEnumerable<string> GetFilesWithDotNet(string path)
    {
        try
        {
            return Directory.EnumerateFiles(path, "*.BDF", SearchOption.AllDirectories);
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private IEnumerable<string> GetFilesWithRobocopy(string path)
    {
        var files = new List<string>();
        try
        {
            var tempDir = Path.Combine(Path.GetTempPath(), $"robocopy_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "robocopy.exe",
                Arguments = $"\"{path}\" \"{tempDir}\" *.BDF /S /L /NJH /NJS /NP /NS /NC /NDL",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(psi);
            if (process != null)
            {
                var output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = line.Trim();
                    if (trimmed.EndsWith(".BDF", StringComparison.OrdinalIgnoreCase))
                    {
                        var fullPath = Path.IsPathRooted(trimmed) ? trimmed : Path.Combine(path, trimmed);
                        files.Add(fullPath);
                    }
                }
            }

            try { Directory.Delete(tempDir, true); } catch { }
        }
        catch { }

        return files;
    }

    private void ExportResults(string filePath)
    {
        try
        {
            var sb = new StringBuilder();

            // Log completo do processamento
            sb.AppendLine("============================================");
            sb.AppendLine("              LOG DE EXECUCAO");
            sb.AppendLine("============================================");
            sb.AppendLine(txtLog.Text);
            sb.AppendLine();

            // Lista de arquivos órfãos
            var orphanList = _orphanFiles.OrderBy(o => o.Tabela).ThenBy(o => o.Handle).ToList();
            sb.AppendLine("============================================");
            sb.AppendLine("     LISTA DE ARQUIVOS ORFAOS");
            sb.AppendLine("============================================");
            sb.AppendLine($"Total: {orphanList.Count} arquivos | Espaco: {FormatSize(_totalOrphanSize)}");
            sb.AppendLine();

            foreach (var orphan in orphanList)
            {
                sb.AppendLine($"Arquivo: {orphan.FileName} | Caminho: {orphan.FilePath} | Sistema: {orphan.Sistema} | Tabela: {orphan.Tabela} | Handle: {orphan.Handle} | Tamanho: {FormatSize(orphan.Size)}");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            Log($"Relatorio exportado: {filePath}", Color.Green);
        }
        catch (Exception ex)
        {
            Log($"Erro ao exportar relatorio: {ex.Message}", Color.Red);
        }
    }

    private static string FormatSize(long bytes)
    {
        if (bytes >= 1073741824) return $"{bytes / 1073741824.0:F2} GB";
        if (bytes >= 1048576) return $"{bytes / 1048576.0:F2} MB";
        if (bytes >= 1024) return $"{bytes / 1024.0:F2} KB";
        return $"{bytes} bytes";
    }

    private void SaveSettings()
    {
        try
        {
            var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
            var lines = new[]
            {
                txtBdocPath.Text,
                txtSystemName.Text,
                txtServer.Text,
                txtDatabase.Text,
                txtUsername.Text,
                cmbSearchMethod.SelectedIndex.ToString(),
                numParallelism.Value.ToString(),
                txtBackupPath.Text
            };
            File.WriteAllLines(settingsPath, lines);
        }
        catch { }
    }

    private void LoadSettings()
    {
        try
        {
            var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
            if (File.Exists(settingsPath))
            {
                var lines = File.ReadAllLines(settingsPath);
                if (lines.Length >= 6)
                {
                    txtBdocPath.Text = lines[0];
                    txtSystemName.Text = lines[1];
                    txtServer.Text = lines[2];
                    txtDatabase.Text = lines[3];
                    txtUsername.Text = lines[4];
                    if (int.TryParse(lines[5], out var index))
                        cmbSearchMethod.SelectedIndex = Math.Min(index, cmbSearchMethod.Items.Count - 1);
                }
                if (lines.Length >= 7 && decimal.TryParse(lines[6], out var parallelism))
                {
                    numParallelism.Value = Math.Clamp(parallelism, numParallelism.Minimum, numParallelism.Maximum);
                }
                if (lines.Length >= 8)
                {
                    txtBackupPath.Text = lines[7];
                }
            }
        }
        catch { }
    }
}

public class OrphanFile
{
    public string Sistema { get; set; } = "";
    public string Tabela { get; set; } = "";
    public int Handle { get; set; }
    public string FileName { get; set; } = "";
    public string FilePath { get; set; } = "";
    public long Size { get; set; }
}

/// <summary>
/// Estrutura leve para armazenar info de arquivo durante processamento paralelo
/// </summary>
public readonly record struct FileHandleInfo(string FilePath, string FileName, int Handle);
