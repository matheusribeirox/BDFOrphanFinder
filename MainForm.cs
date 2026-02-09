using System.Collections.Concurrent;
using System.Data.Common;
using System.Data.SqlClient;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using Oracle.ManagedDataAccess.Client;

namespace BDFOrphanFinder;

#region Win32 API Interop
internal static class NativeMethods
{
    public const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
    public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    // ===== MFT (Master File Table) Access =====
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateFileW(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool DeviceIoControl(
        IntPtr hDevice,
        uint dwIoControlCode,
        ref MFT_ENUM_DATA_V0 lpInBuffer,
        int nInBufferSize,
        IntPtr lpOutBuffer,
        int nOutBufferSize,
        out int lpBytesReturned,
        IntPtr lpOverlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);

    // Constants for MFT access
    public const uint GENERIC_READ = 0x80000000;
    public const uint FILE_SHARE_READ = 0x00000001;
    public const uint FILE_SHARE_WRITE = 0x00000002;
    public const uint OPEN_EXISTING = 3;
    public const uint FSCTL_ENUM_USN_DATA = 0x000900b3;

    [StructLayout(LayoutKind.Sequential)]
    public struct MFT_ENUM_DATA_V0
    {
        public long StartFileReferenceNumber;
        public long LowUsn;
        public long HighUsn;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct USN_RECORD
    {
        public int RecordLength;
        public short MajorVersion;
        public short MinorVersion;
        public long FileReferenceNumber;
        public long ParentFileReferenceNumber;
        public long Usn;
        public long TimeStamp;
        public int Reason;
        public int SourceInfo;
        public int SecurityId;
        public uint FileAttributes;
        public short FileNameLength;
        public short FileNameOffset;
        // FileName follows (variable length)
    }
}
#endregion

#region MFT Scanner
/// <summary>
/// Scanner de alta performance que lê diretamente da MFT (Master File Table) do NTFS
/// Similar ao que o Everything Search faz - indexa milhões de arquivos em segundos
/// REQUER: Execução como Administrador
/// </summary>
internal static class MftScanner
{
    // Limite de memória padrão: 1.5GB (deixa margem para o resto do app)
    private const long DEFAULT_MEMORY_LIMIT_BYTES = 1536L * 1024 * 1024;

    // Tamanho estimado por entrada de diretório em memória (~100 bytes)
    private const int ESTIMATED_DIR_ENTRY_SIZE = 100;

    // Tamanho estimado por arquivo encontrado (~150 bytes)
    private const int ESTIMATED_FILE_ENTRY_SIZE = 150;

    /// <summary>
    /// Escaneia o volume MFT buscando arquivos com a extensão especificada.
    /// OTIMIZAÇÃO: Se basePath for fornecido, apenas diretórios dentro desse caminho serão considerados,
    /// reduzindo drasticamente o uso de memória.
    /// </summary>
    /// <param name="driveLetter">Letra do drive (ex: 'B')</param>
    /// <param name="extension">Extensão do arquivo (ex: ".BDF")</param>
    /// <param name="basePath">Caminho base para filtrar (ex: "B:\bdoc"). Se null, retorna todos.</param>
    /// <param name="progressCallback">Callback de progresso (número de registros processados)</param>
    /// <param name="memoryLimitBytes">Limite de memória em bytes (padrão: 1.5GB)</param>
    public static List<MftFileInfo> ScanVolume(
        char driveLetter,
        string extension,
        string? basePath = null,
        Action<int>? progressCallback = null,
        long memoryLimitBytes = DEFAULT_MEMORY_LIMIT_BYTES)
    {
        var volumePath = $"\\\\.\\{driveLetter}:";

        var handle = NativeMethods.CreateFileW(
            volumePath,
            NativeMethods.GENERIC_READ,
            NativeMethods.FILE_SHARE_READ | NativeMethods.FILE_SHARE_WRITE,
            IntPtr.Zero,
            NativeMethods.OPEN_EXISTING,
            0,
            IntPtr.Zero);

        if (handle == NativeMethods.INVALID_HANDLE_VALUE)
        {
            throw new InvalidOperationException(
                $"Não foi possível abrir o volume {driveLetter}:. Verifique se está executando como Administrador.");
        }

        var files = new List<MftFileInfo>();
        var directories = new Dictionary<long, DirectoryEntry>();

        long estimatedMemory = 0;

        try
        {
            var bufferSize = 64 * 1024; // 64KB buffer
            var buffer = Marshal.AllocHGlobal(bufferSize);

            try
            {
                var mftData = new NativeMethods.MFT_ENUM_DATA_V0
                {
                    StartFileReferenceNumber = 0,
                    LowUsn = 0,
                    HighUsn = long.MaxValue
                };

                int recordCount = 0;
                string extensionUpper = extension.ToUpperInvariant();

                while (NativeMethods.DeviceIoControl(
                    handle,
                    NativeMethods.FSCTL_ENUM_USN_DATA,
                    ref mftData,
                    Marshal.SizeOf(mftData),
                    buffer,
                    bufferSize,
                    out int bytesReturned,
                    IntPtr.Zero))
                {
                    if (bytesReturned <= 8) break;

                    // Primeiro 8 bytes = próximo USN
                    mftData.StartFileReferenceNumber = Marshal.ReadInt64(buffer, 0);

                    int offset = 8;
                    while (offset < bytesReturned)
                    {
                        var record = Marshal.PtrToStructure<NativeMethods.USN_RECORD>(buffer + offset);
                        int recordLength = record.RecordLength;
                        if (recordLength <= 0) break;

                        // Extrair nome do arquivo
                        int nameLength = record.FileNameLength / 2;
                        string fileName = Marshal.PtrToStringUni(
                            buffer + offset + record.FileNameOffset,
                            nameLength) ?? "";

                        if ((record.FileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) != 0)
                        {
                            // É um diretório - armazenar para resolver caminhos depois
                            directories[record.FileReferenceNumber] =
                                new DirectoryEntry(fileName, record.ParentFileReferenceNumber);

                            estimatedMemory += ESTIMATED_DIR_ENTRY_SIZE;
                        }
                        else if (fileName.EndsWith(extensionUpper, StringComparison.OrdinalIgnoreCase))
                        {
                            files.Add(new MftFileInfo
                            {
                                FileReferenceNumber = record.FileReferenceNumber,
                                ParentFileReferenceNumber = record.ParentFileReferenceNumber,
                                FileName = fileName
                            });

                            estimatedMemory += ESTIMATED_FILE_ENTRY_SIZE;
                        }

                        // Verificar limite de memória
                        if (estimatedMemory > memoryLimitBytes)
                        {
                            // Compactar diretórios antes de lançar erro
                            CompactDirectories(directories, files);
                            estimatedMemory = (long)directories.Count * ESTIMATED_DIR_ENTRY_SIZE +
                                            (long)files.Count * ESTIMATED_FILE_ENTRY_SIZE;

                            if (estimatedMemory > memoryLimitBytes)
                            {
                                throw new OutOfMemoryException(
                                    $"Limite de memória atingido ({memoryLimitBytes / (1024 * 1024)} MB). " +
                                    $"Dirs: {directories.Count:N0}, Arquivos: {files.Count:N0}");
                            }
                        }

                        recordCount++;
                        if (recordCount % 100000 == 0)
                        {
                            progressCallback?.Invoke(recordCount);
                        }

                        offset += recordLength;
                    }
                }

                // Resolver caminhos completos e filtrar por basePath
                progressCallback?.Invoke(-1); // Sinaliza início da resolução de paths

                var filteredFiles = new List<MftFileInfo>();
                int resolvedCount = 0;
                string? basePathUpper = basePath?.ToUpperInvariant().TrimEnd('\\');

                foreach (var file in files)
                {
                    var fullPath = $"{driveLetter}:{BuildPath(file.FileName, file.ParentFileReferenceNumber, directories)}";
                    file.FullPath = fullPath;
                    resolvedCount++;

                    // Se temos basePath, filtrar apenas arquivos dentro dele
                    if (basePathUpper != null)
                    {
                        if (fullPath.ToUpperInvariant().StartsWith(basePathUpper + "\\"))
                        {
                            filteredFiles.Add(file);
                        }
                    }
                    else
                    {
                        filteredFiles.Add(file);
                    }

                    if (resolvedCount % 10000 == 0)
                    {
                        progressCallback?.Invoke(-resolvedCount);
                    }
                }

                // Liberar memória dos diretórios (não mais necessários)
                directories.Clear();
                directories.TrimExcess();

                // Retornar apenas arquivos filtrados
                files = filteredFiles;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        finally
        {
            NativeMethods.CloseHandle(handle);
        }

        return files;
    }

    private static void CompactDirectories(Dictionary<long, DirectoryEntry> directories, List<MftFileInfo> files)
    {
        // Identificar diretórios que são ancestrais dos arquivos encontrados
        var neededDirs = new HashSet<long>();

        foreach (var file in files)
        {
            var parentRef = file.ParentFileReferenceNumber;
            while (parentRef != 0 && directories.ContainsKey(parentRef))
            {
                if (!neededDirs.Add(parentRef)) break; // Já processado
                parentRef = directories[parentRef].ParentRef;
            }
        }

        // Remover diretórios não necessários
        var keysToRemove = directories.Keys.Where(k => !neededDirs.Contains(k)).ToList();
        foreach (var key in keysToRemove)
        {
            directories.Remove(key);
        }

        directories.TrimExcess();
    }

    private static string BuildPath(string fileName, long parentRef, Dictionary<long, DirectoryEntry> directories)
    {
        var parts = new Stack<string>();
        parts.Push(fileName);

        var currentParent = parentRef;
        int maxDepth = 100; // Prevenir loops infinitos

        while (currentParent != 0 && directories.TryGetValue(currentParent, out var parent) && maxDepth-- > 0)
        {
            parts.Push(parent.Name);
            currentParent = parent.ParentRef;
        }

        return "\\" + string.Join("\\", parts);
    }

    // Estrutura compacta para diretórios (economiza memória)
    private readonly record struct DirectoryEntry(string Name, long ParentRef);
}

internal class MftFileInfo
{
    public long FileReferenceNumber { get; set; }
    public long ParentFileReferenceNumber { get; set; }
    public string FileName { get; set; } = "";
    public string FullPath { get; set; } = "";
}
#endregion

public partial class MainForm : Form
{
    private TextBox txtBdocPath = null!;
    private TextBox txtAppServer = null!;
    private ComboBox cmbSistema = null!;
    private ComboBox cmbSearchMethod = null!;
    private NumericUpDown numParallelism = null!;
    private TextBox txtBackupPath = null!;
    private CheckBox chkRemoveAfterBackup = null!;
    private Button btnBrowse = null!;
    private Button btnConectar = null!;
    private Button btnBrowseBackup = null!;
    private Button btnStart = null!;
    private Button btnCancel = null!;
    private ProgressBar progressBar = null!;
    private Label lblStatus = null!;
    private Label lblProgress = null!;
    private RichTextBox txtLog = null!;

    // Constantes
    private const int PROGRESS_UPDATE_INTERVAL = 100;
    private const int DB_COMMAND_TIMEOUT = 120;
    private const int ORACLE_BATCH_SIZE = 1000;
    private const int SQLSERVER_BATCH_SIZE = 2000;
    private const int TELNET_TIMEOUT_MS = 10000;

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

    // Tipo de banco detectado automaticamente via connection string
    private bool _isOracle = false;
    private bool IsOracle => _isOracle;

    // Diretórios BDOC obtidos do BSERVER (PARDIR + SECDIR)
    private List<string> _bdocPaths = new();

    public MainForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Identificador de Arquivos BDF Órfãos";
        this.Size = new Size(900, 750);
        this.MinimumSize = new Size(800, 650);
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
            Text = "Configurações",
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
            RowCount = 6,
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

        // Row 0: BDOC Path + Servidor de Aplicação
        configPanel.Controls.Add(new Label { Text = "Diretório BDOC:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 0);
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
        configPanel.Controls.Add(new Label { Text = "Serv. Aplicação:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 0);
        txtAppServer = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtAppServer, 4, 0);
        btnConectar = new Button
        {
            Text = "Conectar", Width = 70, Height = 26,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 8.5f),
            Cursor = Cursors.Hand
        };
        btnConectar.FlatAppearance.BorderSize = 0;
        btnConectar.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 180);
        btnConectar.Click += BtnConectar_Click;
        configPanel.Controls.Add(btnConectar, 5, 0);

        // Row 1: Sistema (ComboBox populado automaticamente)
        configPanel.Controls.Add(new Label { Text = "Sistema:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 1);
        cmbSistema = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Standard,
            BackColor = Color.White,
            Font = textBoxFont
        };
        configPanel.Controls.Add(cmbSistema, 1, 1);
        configPanel.SetColumnSpan(cmbSistema, 4);

        // Row 2: Search Method + Parallelism
        configPanel.Controls.Add(new Label { Text = "Método de Busca:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 2);
        cmbSearchMethod = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Standard,
            BackColor = Color.White,
            Font = textBoxFont
        };
        // 0 = MFT (Master File Table), 1 = .NET EnumerateFiles
        cmbSearchMethod.Items.AddRange(new object[] { "MFT (Master File Table)", ".NET EnumerateFiles" });
        cmbSearchMethod.SelectedIndex = 0;
        configPanel.Controls.Add(cmbSearchMethod, 1, 2);

        configPanel.Controls.Add(new Label { Text = "Threads:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 2);
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
        configPanel.Controls.Add(numParallelism, 4, 2);

        // Row 3: Backup Path
        configPanel.Controls.Add(new Label { Text = "Caminho Backup BDF:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 3);
        txtBackupPath = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtBackupPath, 1, 3);
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
        configPanel.Controls.Add(btnBrowseBackup, 2, 3);

        // Checkbox para remover após backup
        chkRemoveAfterBackup = new CheckBox
        {
            Text = "Remover órfãos",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(10, 6, 3, 0),
            ForeColor = Color.FromArgb(180, 0, 0),
            Font = labelFont,
            Cursor = Cursors.Hand
        };
        configPanel.Controls.Add(chkRemoveAfterBackup, 3, 3);
        configPanel.SetColumnSpan(chkRemoveAfterBackup, 2);

        // Row 4: Buttons
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

        configPanel.Controls.Add(buttonPanel, 1, 4);
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
            Description = "Selecione o diretório base do BDOC",
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
            Description = "Selecione o diretório de backup dos BDF órfãos",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtBackupPath.Text = dialog.SelectedPath;
        }
    }

    // ===== Comunicação Telnet com Servidor de Aplicação =====

    private async Task<string> SendTelnetCommandAsync(string server, int port, string command, CancellationToken ct)
    {
        using var client = new TcpClient();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(TELNET_TIMEOUT_MS);

        await client.ConnectAsync(server, port, timeoutCts.Token);
        using var stream = client.GetStream();

        var data = Encoding.ASCII.GetBytes(command + "\r\n");
        await stream.WriteAsync(data, timeoutCts.Token);

        // Ler resposta
        var buffer = new byte[8192];
        var response = new StringBuilder();

        while (true)
        {
            var bytesRead = await stream.ReadAsync(buffer, timeoutCts.Token);
            if (bytesRead == 0) break;
            response.Append(Encoding.ASCII.GetString(buffer, 0, bytesRead));
            if (response.ToString().Contains("\n")) break;
        }
        return response.ToString().Trim();
    }

    private async Task<string> SendCommandOnStreamAsync(NetworkStream stream, string command, CancellationToken ct)
    {
        var data = Encoding.ASCII.GetBytes(command + "\r\n");
        await stream.WriteAsync(data, ct);

        var buffer = new byte[8192];
        var response = new StringBuilder();

        while (true)
        {
            var bytesRead = await stream.ReadAsync(buffer, ct);
            if (bytesRead == 0) break;
            response.Append(Encoding.ASCII.GetString(buffer, 0, bytesRead));
            if (response.ToString().Contains("\n")) break;
        }
        return response.ToString().Trim();
    }

    private async Task SendAndExpectAsync(NetworkStream stream, string command, string expectedPrefix, CancellationToken ct)
    {
        var response = await SendCommandOnStreamAsync(stream, command, ct);
        if (!response.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Resposta inesperada do servidor.\nComando: {command}\nEsperado: {expectedPrefix}\nRecebido: {response}");
        }
    }

    private async Task<string> GetBserverConnectionStringAsync(string server, CancellationToken ct)
    {
        var response = await SendTelnetCommandAsync(server, 5331, "getssdbadonetconnectionstring", ct);
        // Remover prefixo "+ "
        if (response.StartsWith("+ "))
            return response.Substring(2);
        if (response.StartsWith("+"))
            return response.Substring(1).TrimStart();
        return response;
    }

    private async Task<string> GetSystemConnectionStringAsync(string server, string systemName, CancellationToken ct)
    {
        using var client = new TcpClient();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(TELNET_TIMEOUT_MS);

        await client.ConnectAsync(server, 5337, timeoutCts.Token);
        using var stream = client.GetStream();

        // 1. Autenticar
        await SendAndExpectAsync(stream, "user internal benner", "+", timeoutCts.Token);

        // 2. Selecionar sistema
        await SendAndExpectAsync(stream, $"selectsystem {systemName}", "+", timeoutCts.Token);

        // 3. Obter connection string
        var response = await SendCommandOnStreamAsync(stream, "getadonetconnectionstring", timeoutCts.Token);

        // Remover prefixo "+ "
        if (response.StartsWith("+ "))
            return response.Substring(2);
        if (response.StartsWith("+"))
            return response.Substring(1).TrimStart();
        return response;
    }

    private (bool isOracle, string connectionString) ParseReceivedConnectionString(string raw)
    {
        bool isOracle = raw.Contains("(DESCRIPTION=", StringComparison.OrdinalIgnoreCase);
        return (isOracle, raw);
    }

    private async Task<List<string>> QueryAvailableSystemsAsync(string connectionString, bool isOracle, CancellationToken ct)
    {
        var systems = new List<string>();

        DbConnection connection = isOracle
            ? new OracleConnection(connectionString)
            : new SqlConnection(connectionString);

        using (connection)
        {
            await connection.OpenAsync(ct);

            using var cmd = connection.CreateCommand();
            cmd.CommandText = isOracle
                ? "SELECT NAME FROM \"SYS_SYSTEMS\" WHERE LICINFO IS NOT NULL"
                : "SELECT NAME FROM [SYS_SYSTEMS] WITH (NOLOCK) WHERE LICINFO IS NOT NULL";
            cmd.CommandTimeout = DB_COMMAND_TIMEOUT;

            using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                if (!reader.IsDBNull(0))
                    systems.Add(reader.GetString(0));
            }
        }

        return systems;
    }

    private async Task<List<string>> QueryBdocPathsAsync(string connectionString, bool isOracle, CancellationToken ct)
    {
        var paths = new List<string>();

        DbConnection connection = isOracle
            ? new OracleConnection(connectionString)
            : new SqlConnection(connectionString);

        using (connection)
        {
            await connection.OpenAsync(ct);

            using var cmd = connection.CreateCommand();
            cmd.CommandText = isOracle
                ? "SELECT NAME, DATA FROM \"SER_SERVICEPARAMS\" WHERE NAME IN ('SECDIR','PARDIR') AND SERVICE = 4"
                : "SELECT NAME, DATA FROM [SER_SERVICEPARAMS] WITH (NOLOCK) WHERE NAME IN ('SECDIR','PARDIR') AND SERVICE = 4";
            cmd.CommandTimeout = DB_COMMAND_TIMEOUT;

            using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                if (!reader.IsDBNull(1))
                {
                    var data = reader.GetString(1);
                    // DATA pode conter múltiplos caminhos separados por ;
                    foreach (var path in data.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    {
                        var trimmed = path.Trim().TrimEnd('\\');
                        if (!string.IsNullOrWhiteSpace(trimmed))
                            paths.Add(trimmed);
                    }
                }
            }
        }

        var distinct = paths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // Tentar converter caminhos UNC para locais quando o servidor é a máquina atual
        var result = new List<string>();
        foreach (var p in distinct)
        {
            var local = TryResolveUncToLocal(p);
            result.Add(local ?? p);
        }

        return result;
    }

    /// <summary>
    /// Se o caminho UNC aponta para a máquina atual, resolve para o caminho local do compartilhamento.
    /// Ex: \\MAQUINA\BDOC\subdir -> D:\BDOC\subdir (se o share BDOC aponta para D:\BDOC)
    /// </summary>
    private string? TryResolveUncToLocal(string uncPath)
    {
        if (!uncPath.StartsWith("\\\\")) return null;

        // Extrair servidor e share do caminho UNC
        var parts = uncPath.TrimStart('\\').Split('\\', 3);
        if (parts.Length < 2) return null;

        var server = parts[0];
        var shareName = parts[1];
        var remainder = parts.Length > 2 ? parts[2] : "";

        // Verificar se o servidor UNC é a máquina atual
        var machineName = Environment.MachineName;
        if (!server.Equals(machineName, StringComparison.OrdinalIgnoreCase))
            return null;

        // Consultar o caminho local do compartilhamento via WMI
        try
        {
            using var searcher = new System.Management.ManagementObjectSearcher(
                $"SELECT Path FROM Win32_Share WHERE Name = '{shareName.Replace("'", "''")}'");
            foreach (var share in searcher.Get())
            {
                var localSharePath = share["Path"]?.ToString();
                if (!string.IsNullOrWhiteSpace(localSharePath))
                {
                    var localFull = string.IsNullOrEmpty(remainder)
                        ? localSharePath.TrimEnd('\\')
                        : Path.Combine(localSharePath, remainder);
                    return localFull;
                }
            }
        }
        catch
        {
            // WMI não disponível ou sem permissão – manter UNC
        }

        return null;
    }

    private async void BtnConectar_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtAppServer.Text))
        {
            MessageBox.Show("Informe o servidor de aplicação.", "Validação",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        btnConectar.Enabled = false;
        cmbSistema.Items.Clear();
        _bdocPaths.Clear();

        try
        {
            Log("Conectando ao servidor de aplicação...", Color.Cyan);
            lblStatus.Text = "Conectando...";

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

            // 1. Obter connection string do BSERVER via porta 5331
            var bserverCs = await GetBserverConnectionStringAsync(txtAppServer.Text, cts.Token);
            Log("Connection string BSERVER obtida com sucesso!", Color.Green);

            // 2. Detectar tipo de banco do BSERVER
            var (isOracle, cs) = ParseReceivedConnectionString(bserverCs);
            var dbType = isOracle ? "Oracle" : "SQL Server";
            Log($"Tipo de banco BSERVER: {dbType}");

            // 3. Garantir parâmetros adicionais na connection string
            var adjustedCs = AdjustConnectionString(cs, isOracle, (int)numParallelism.Value);

            // 4. Consultar diretórios BDOC (PARDIR + SECDIR)
            Log("Consultando diretórios BDOC...");
            _bdocPaths = await QueryBdocPathsAsync(adjustedCs, isOracle, cts.Token);
            foreach (var p in _bdocPaths)
                Log($"  [BDOC] {p}", Color.Gray);

            if (_bdocPaths.Count > 0)
            {
                // Preencher campo BDOC com o primeiro diretório local encontrado (ou o primeiro disponível)
                var localPath = _bdocPaths.FirstOrDefault(p => !p.StartsWith("\\\\")) ?? _bdocPaths[0];
                txtBdocPath.Text = localPath;
                Log($"Diretório BDOC selecionado: {localPath}", Color.Green);
            }

            // 5. Consultar sistemas disponíveis
            Log("Consultando sistemas disponíveis...");
            var systems = await QueryAvailableSystemsAsync(adjustedCs, isOracle, cts.Token);

            // 6. Popular ComboBox
            foreach (var sys in systems.OrderBy(s => s))
                cmbSistema.Items.Add(sys);

            if (cmbSistema.Items.Count > 0)
                cmbSistema.SelectedIndex = 0;

            Log($"{systems.Count} sistema(s) encontrado(s).", Color.Green);
            lblStatus.Text = $"{systems.Count} sistema(s) carregado(s)";
        }
        catch (Exception ex)
        {
            Log($"ERRO ao conectar: {ex.Message}", Color.Red);
            lblStatus.Text = "Erro na conexão";
            MessageBox.Show($"Erro ao conectar ao servidor de aplicação:\n{ex.Message}", "Erro",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnConectar.Enabled = true;
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
            Log("Operação cancelada pelo usuário.", Color.Yellow);
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
            MessageBox.Show("Diretório BDOC inválido ou não existe.", "Validação",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (string.IsNullOrWhiteSpace(txtAppServer.Text))
        {
            MessageBox.Show("Servidor de aplicação é obrigatório.", "Validação",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (cmbSistema.SelectedIndex < 0 || string.IsNullOrWhiteSpace(cmbSistema.Text))
        {
            MessageBox.Show("Selecione um sistema. Clique em 'Conectar' para carregar os sistemas disponíveis.", "Validação",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (!string.IsNullOrWhiteSpace(txtBackupPath.Text) && !Directory.Exists(txtBackupPath.Text))
        {
            MessageBox.Show("Diretório de backup inválido ou não existe.", "Validação",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        // Se checkbox de remover está marcado sem backup, pede confirmação
        if (chkRemoveAfterBackup.Checked && string.IsNullOrWhiteSpace(txtBackupPath.Text))
        {
            var result = MessageBox.Show(
                "A opção de remover arquivos órfãos está marcada, porém nenhum caminho de backup foi configurado.\n\n" +
                "Os arquivos serão removidos permanentemente sem backup!\n\n" +
                "Deseja continuar mesmo assim?",
                "Atenção - Sem Backup",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
                return false;
        }

        return true;
    }

    private void SetUIState(bool running)
    {
        btnStart.Enabled = !running;
        btnCancel.Enabled = running;
        btnBrowse.Enabled = !running;
        btnConectar.Enabled = !running;
        btnBrowseBackup.Enabled = !running;
        txtBdocPath.Enabled = !running;
        txtAppServer.Enabled = !running;
        txtBackupPath.Enabled = !running;
        cmbSistema.Enabled = !running;
        cmbSearchMethod.Enabled = !running;
        numParallelism.Enabled = !running;

        // Ajustar cor do botão Iniciar conforme estado
        btnStart.BackColor = running
            ? Color.FromArgb(180, 180, 180)
            : Color.FromArgb(0, 120, 212);

        progressBar.Value = 0;
        lblProgress.Text = running ? "0%" : "";
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

    // ===== Métodos auxiliares de abstração de banco =====

    private string AdjustConnectionString(string connectionString, bool isOracle, int maxParallelism)
    {
        // Adicionar Max Pool Size se não estiver presente
        if (!connectionString.Contains("Max Pool Size", StringComparison.OrdinalIgnoreCase))
        {
            connectionString = connectionString.TrimEnd(';') + $";Max Pool Size={maxParallelism + 5};";
        }

        // Para SQL Server, garantir TrustServerCertificate
        if (!isOracle && !connectionString.Contains("TrustServerCertificate", StringComparison.OrdinalIgnoreCase))
        {
            connectionString = connectionString.TrimEnd(';') + ";TrustServerCertificate=True;";
        }

        return connectionString;
    }

    private DbConnection CreateConnection(string connectionString)
    {
        if (IsOracle)
            return new OracleConnection(connectionString);
        return new SqlConnection(connectionString);
    }

    /// <summary>
    /// Valida se o nome da tabela contém apenas caracteres seguros para uso em SQL
    /// </summary>
    private static bool IsValidTableName(string tableName)
    {
        // Aceitar apenas letras, números, underscore e cifrão (padrão BDOC: DO_ADMISSAODOCUMENTOS)
        return !string.IsNullOrWhiteSpace(tableName) &&
               tableName.All(c => char.IsLetterOrDigit(c) || c == '_' || c == '$');
    }

    private string BuildHandleQuery(string tableName, string? inClause = null)
    {
        if (!IsValidTableName(tableName))
            throw new ArgumentException($"Nome de tabela inválido: {tableName}");

        if (IsOracle)
        {
            var q = $"SELECT DISTINCT HANDLE FROM \"{tableName.ToUpperInvariant()}\"";
            return inClause != null ? $"{q} WHERE HANDLE IN ({inClause})" : q;
        }
        else
        {
            var q = $"SELECT DISTINCT HANDLE FROM [{tableName}] WITH (NOLOCK)";
            return inClause != null ? $"{q} WHERE HANDLE IN ({inClause})" : q;
        }
    }

    private int GetBatchSize()
    {
        return IsOracle ? ORACLE_BATCH_SIZE : SQLSERVER_BATCH_SIZE;
    }

    private async Task ProcessOrphanFilesAsync(CancellationToken ct)
    {
        var systemName = cmbSistema.Text;
        var maxParallelism = (int)numParallelism.Value;
        var searchMethod = cmbSearchMethod.SelectedIndex;

        // Determinar diretórios BDOC a usar
        var bdocDirs = new List<string>();
        if (!string.IsNullOrWhiteSpace(txtBdocPath.Text))
        {
            bdocDirs.Add(txtBdocPath.Text.TrimEnd('\\'));
        }
        // Adicionar diretórios do BSERVER que não estejam já na lista
        foreach (var p in _bdocPaths)
        {
            if (!bdocDirs.Any(d => d.Equals(p, StringComparison.OrdinalIgnoreCase)))
                bdocDirs.Add(p);
        }

        Log("Iniciando processamento...", Color.Cyan);
        Log($"Diretórios BDOC: {string.Join("; ", bdocDirs)}");
        Log($"Sistema: {systemName}");
        Log($"Servidor de aplicação: {txtAppServer.Text}");
        Log($"Método: {cmbSearchMethod.SelectedItem}");
        Log($"Threads: {maxParallelism}");
        Log("");

        var startTime = DateTime.Now;

        // Find system paths - procurar em dois níveis:
        // 1. Direto na raiz do BDOC: BDOC\SISTEMA\
        // 2. Um nível abaixo: BDOC\subdir\SISTEMA\
        Log("Procurando caminhos do sistema...");
        var systemPaths = new List<string>();

        foreach (var bdocPath in bdocDirs)
        {
            if (!Directory.Exists(bdocPath))
            {
                Log($"  [AVISO] Diretório não acessível: {bdocPath}", Color.Yellow);
                continue;
            }

            // Nível 1: BDOC\SISTEMA\ (direto na raiz)
            var directPath = Path.Combine(bdocPath, systemName);
            if (Directory.Exists(directPath))
            {
                systemPaths.Add(directPath);
                Log($"  [ENCONTRADO] {directPath}", Color.Green);
            }

            // Nível 2: BDOC\subdir\SISTEMA\ (um nível abaixo)
            try
            {
                foreach (var subdir in Directory.EnumerateDirectories(bdocPath))
                {
                    var subdirName = Path.GetFileName(subdir);
                    // Evitar duplicata se o subdir for o próprio sistema (já encontrado acima)
                    if (subdirName.Equals(systemName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var systemPath = Path.Combine(subdir, systemName);
                    if (Directory.Exists(systemPath))
                    {
                        systemPaths.Add(systemPath);
                        Log($"  [ENCONTRADO] {systemPath}", Color.Green);
                    }
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (DirectoryNotFoundException) { }
        }

        if (systemPaths.Count == 0)
        {
            Log($"ERRO: Sistema '{systemName}' não encontrado em nenhum diretório BDOC.", Color.Red);
            return;
        }

        // Usar o primeiro BDOC como referência para caminhos relativos (backup)
        var bdocPath_ref = bdocDirs[0];

        // Obter connection string do sistema via telnet porta 5337
        Log("Obtendo connection string do sistema via servidor de aplicação...", Color.Cyan);
        var rawCs = await GetSystemConnectionStringAsync(txtAppServer.Text, systemName, ct);

        // Detectar tipo de banco
        var (isOracle, cs) = ParseReceivedConnectionString(rawCs);
        _isOracle = isOracle;
        var dbType = IsOracle ? "Oracle" : "SQL Server";
        Log($"Tipo de banco detectado: {dbType}", Color.Green);

        // Ajustar connection string
        var connectionString = AdjustConnectionString(cs, isOracle, maxParallelism);

        // Test connection
        Log($"Testando conexão com banco de dados {dbType}...");
        using (var testConn = CreateConnection(connectionString))
        {
            await testConn.OpenAsync(ct);
            Log("Conexão estabelecida com sucesso!", Color.Green);
        }

        // Use strategy based on selection
        // 0 = MFT (Master File Table) - requer Admin
        // 1 = .NET EnumerateFiles
        if (searchMethod == 0) // MFT
        {
            // MFT não suporta caminhos de rede (UNC)
            if (bdocPath_ref.StartsWith("\\\\"))
            {
                Log("ERRO: O método MFT não suporta caminhos de rede (UNC).", Color.Red);
                Log("Use o método '.NET EnumerateFiles' para caminhos de rede.", Color.Yellow);
                return;
            }

            // Verificar se está rodando como Administrador
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            if (!principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator))
            {
                Log("ERRO: O método MFT requer execução como Administrador.", Color.Red);
                Log("Execute o programa como Administrador ou use o método '.NET EnumerateFiles'.", Color.Yellow);
                return;
            }

            await ProcessWithMftStrategyAsync(systemPaths, systemName, connectionString, maxParallelism, ct);
        }
        else // .NET EnumerateFiles (Híbrido Paralelo)
        {
            await ProcessWithHybridStrategyAsync(systemPaths, systemName, connectionString, maxParallelism, ct);
        }

        // ================================================================
        // FASE 3: Backup dos arquivos órfãos (se configurado)
        // ================================================================
        var backupPath = txtBackupPath.Text;
        var shouldRemoveOrphans = false;
        var filesRemoved = 0;
        var removeErrors = 0;

        var backupCancelled = false;

        if (!string.IsNullOrWhiteSpace(backupPath) && !_orphanFiles.IsEmpty)
        {
            // Se checkbox de remoção está marcado, pedir confirmação
            if (chkRemoveAfterBackup.Checked)
            {
                var confirmResult = MessageBox.Show(
                    $"Foram encontrados {_orphanFiles.Count} arquivos órfãos ({FormatSize(_totalOrphanSize)}).\n\n" +
                    $"Os arquivos serão copiados para:\n{backupPath}\n\n" +
                    "ATENÇÃO: Após o backup, os arquivos originais serão REMOVIDOS do BDOC!\n\n" +
                    "Deseja continuar com o backup e remoção?",
                    "Confirmação de Remoção",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (confirmResult != DialogResult.Yes)
                {
                    Log("Operação de backup/remoção cancelada pelo usuário.", Color.Yellow);
                    backupCancelled = true;
                }
                else
                {
                    shouldRemoveOrphans = true;
                }
            }

            if (!backupCancelled)
            {
            Log("");
            Log("=== FASE 3: Backup dos arquivos órfãos ===", Color.Cyan);
            var phase3Start = DateTime.Now;

            var orphanList = _orphanFiles.ToList();
            var copied = 0;
            var errors = 0;
            var totalOrphans = orphanList.Count;

            // Lista de arquivos copiados com sucesso (para remoção posterior)
            var successfullyCopied = new ConcurrentBag<string>();

            Log($"Copiando {totalOrphans} arquivos para: {backupPath}");

            await Parallel.ForEachAsync(orphanList,
                new ParallelOptions { MaxDegreeOfParallelism = (int)numParallelism.Value, CancellationToken = ct },
                async (orphan, token) =>
                {
                    try
                    {
                        // Encontrar o BDOC base correto para este arquivo
                        var baseBdoc = bdocDirs.FirstOrDefault(d => orphan.FilePath.StartsWith(d + "\\", StringComparison.OrdinalIgnoreCase)) ?? bdocPath_ref;
                        var relativePath = Path.GetRelativePath(baseBdoc, orphan.FilePath);
                        var destPath = Path.Combine(backupPath, relativePath);
                        var destDir = Path.GetDirectoryName(destPath)!;

                        Directory.CreateDirectory(destDir);

                        // Se arquivo já existe no destino (re-execução), sobrescrever
                        File.Copy(orphan.FilePath, destPath, overwrite: true);

                        // Marcar como copiado com sucesso
                        successfullyCopied.Add(orphan.FilePath);

                        var done = Interlocked.Increment(ref copied);
                        if (done % PROGRESS_UPDATE_INTERVAL == 0 || done == totalOrphans)
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

            Log($"Backup concluído: {copied} copiados, {errors} erros", copied > 0 ? Color.Green : Color.Yellow);
            Log($"Fase 3 concluída em: {_phase3Duration:mm\\:ss\\.fff}", Color.Green);

            // ================================================================
            // FASE 4: Remoção dos arquivos órfãos do BDOC (se solicitado)
            // ================================================================
            if (shouldRemoveOrphans && !successfullyCopied.IsEmpty)
            {
                Log("");
                Log("=== FASE 4: Remoção dos arquivos órfãos do BDOC ===", Color.Cyan);
                var phase4Start = DateTime.Now;

                var filesToRemove = successfullyCopied.ToList();
                var totalToRemove = filesToRemove.Count;
                var dirsRemoved = 0;

                Log($"Removendo {totalToRemove} arquivos do BDOC...");

                // Passo 1: Deletar arquivos em paralelo (sem tocar em pastas)
                await Parallel.ForEachAsync(filesToRemove,
                    new ParallelOptions { MaxDegreeOfParallelism = (int)numParallelism.Value, CancellationToken = ct },
                    async (filePath, token) =>
                    {
                        try
                        {
                            File.Delete(filePath);
                            var done = Interlocked.Increment(ref filesRemoved);
                            if (done % PROGRESS_UPDATE_INTERVAL == 0 || done == totalToRemove)
                            {
                                UpdateProgress(done, totalToRemove, $"Fase 4: Removendo ({done}/{totalToRemove} arquivos)");
                            }
                        }
                        catch (Exception ex)
                        {
                            Interlocked.Increment(ref removeErrors);
                            Log($"  [ERRO REMOÇÃO] {Path.GetFileName(filePath)}: {ex.Message}", Color.Yellow);
                        }

                        await Task.CompletedTask;
                    });

                // Passo 2: Limpar pastas vazias sequencialmente (evita race condition)
                var normalizedSystemPaths = systemPaths.Select(p => p.TrimEnd('\\').ToUpperInvariant()).ToHashSet();

                // Coletar pastas candidatas e ordenar do mais profundo para o mais raso
                var candidateDirs = filesToRemove
                    .Select(f => Path.GetDirectoryName(f))
                    .Where(d => !string.IsNullOrEmpty(d))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(d => d!.Length)
                    .ToList();

                foreach (var dir in candidateDirs)
                {
                    ct.ThrowIfCancellationRequested();
                    var currentDir = dir;

                    while (!string.IsNullOrEmpty(currentDir))
                    {
                        var normalizedCurrentDir = currentDir.TrimEnd('\\').ToUpperInvariant();

                        // Parar se chegou no diretório da tabela ou acima
                        var isTableOrAbove = normalizedSystemPaths.Any(sp =>
                            normalizedCurrentDir.Equals(sp) ||
                            sp.StartsWith(normalizedCurrentDir + "\\") ||
                            normalizedCurrentDir.Equals(sp) ||
                            (normalizedCurrentDir.StartsWith(sp + "\\") &&
                             normalizedCurrentDir.Substring(sp.Length + 1).IndexOf('\\') == -1));

                        if (isTableOrAbove)
                            break;

                        try
                        {
                            if (Directory.Exists(currentDir) && !Directory.EnumerateFileSystemEntries(currentDir).Any())
                            {
                                Directory.Delete(currentDir);
                                dirsRemoved++;
                                currentDir = Path.GetDirectoryName(currentDir);
                            }
                            else
                            {
                                break;
                            }
                        }
                        catch
                        {
                            break;
                        }
                    }
                }

                var phase4Duration = DateTime.Now - phase4Start;
                Log($"Remoção concluída: {filesRemoved} arquivos, {dirsRemoved} pastas vazias, {removeErrors} erros", filesRemoved > 0 ? Color.Green : Color.Yellow);
                Log($"Fase 4 concluída em: {phase4Duration:mm\\:ss\\.fff}", Color.Green);

                // Gerar log de auditoria da remoção
                try
                {
                    var auditFileName = $"auditoria_remocao_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    var auditPath = Path.Combine(Environment.CurrentDirectory, auditFileName);
                    var auditSb = new StringBuilder();
                    auditSb.AppendLine("============================================");
                    auditSb.AppendLine("       AUDITORIA DE REMOÇÃO DE ÓRFÃOS");
                    auditSb.AppendLine("============================================");
                    auditSb.AppendLine($"Data/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    auditSb.AppendLine($"Usuário: {Environment.UserName}");
                    auditSb.AppendLine($"Máquina: {Environment.MachineName}");
                    auditSb.AppendLine($"BDOC: {string.Join("; ", bdocDirs)}");
                    auditSb.AppendLine($"Backup: {backupPath}");
                    auditSb.AppendLine($"Arquivos removidos: {filesRemoved}");
                    auditSb.AppendLine($"Pastas removidas: {dirsRemoved}");
                    auditSb.AppendLine($"Erros: {removeErrors}");
                    auditSb.AppendLine();
                    auditSb.AppendLine("ARQUIVOS REMOVIDOS:");
                    foreach (var f in filesToRemove)
                    {
                        auditSb.AppendLine(f);
                    }
                    File.WriteAllText(auditPath, auditSb.ToString(), Encoding.UTF8);
                    Log($"Log de auditoria: {auditPath}", Color.Green);
                }
                catch (Exception ex)
                {
                    Log($"Erro ao gerar log de auditoria: {ex.Message}", Color.Yellow);
                }
            }
            } // fim if (!backupCancelled)
        }

        // Final report
        var elapsed = DateTime.Now - startTime;
        var orphanCount = _orphanFiles.Count;

        Log("", Color.White);
        Log("============================================", Color.Cyan);
        Log("           RELATÓRIO FINAL", Color.Cyan);
        Log("============================================", Color.Cyan);
        Log($"Banco de dados: {dbType}");
        Log($"Total de arquivos analisados: {_totalFilesAnalyzed}");
        Log($"Arquivos COM registro na base: {_filesWithRecord}", Color.Green);
        Log($"Arquivos ÓRFÃOS (sem registro): {orphanCount}", orphanCount > 0 ? Color.Red : Color.Green);
        Log($"Espaço ocupado por órfãos: {FormatSize(_totalOrphanSize)}", Color.Yellow);
        if (_totalFilesCopied > 0 || _totalCopyErrors > 0)
        {
            Log($"Backup: {_totalFilesCopied} copiados, {_totalCopyErrors} erros", Color.Cyan);
        }
        if (filesRemoved > 0 || removeErrors > 0)
        {
            Log($"Remoção: {filesRemoved} removidos, {removeErrors} erros", Color.Cyan);
        }
        Log("");
        Log("--- MÉTRICAS DE PERFORMANCE ---", Color.Cyan);
        Log($"Fase 1 (Enumeração de arquivos): {_phase1Duration:hh\\:mm\\:ss\\.fff}");
        Log($"Fase 2 (Validação no banco): {_phase2Duration:hh\\:mm\\:ss\\.fff}");
        if (_phase3Duration > TimeSpan.Zero)
        {
            Log($"Fase 3 (Backup de órfãos): {_phase3Duration:hh\\:mm\\:ss\\.fff}");
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
    /// ESTRATÉGIA MFT - ULTRA PERFORMANCE
    /// Lê diretamente da Master File Table do NTFS (como o Everything faz)
    /// REQUER: Execução como Administrador
    /// </summary>
    private async Task ProcessWithMftStrategyAsync(
        List<string> systemPaths,
        string systemName,
        string connectionString,
        int maxParallelism,
        CancellationToken ct)
    {
        // ================================================================
        // FASE 1: Leitura direta da MFT
        // ================================================================
        Log("");
        Log("=== FASE 1: Leitura direta da MFT (Master File Table) ===", Color.Cyan);
        Log("NOTA: Este método requer execução como Administrador", Color.Yellow);
        var phase1Start = DateTime.Now;

        // Descobrir drive letter do BDOC path
        var bdocPath = txtBdocPath.Text;
        var driveLetter = Path.GetPathRoot(bdocPath)?[0] ?? 'C';

        Log($"Escaneando volume {driveLetter}: ...");

        List<MftFileInfo> allBdfFiles;
        try
        {
            // Limite de memória: 1.5GB para o scanner MFT
            const long memoryLimit = 1536L * 1024 * 1024;
            Log($"Limite de memória configurado: {memoryLimit / (1024 * 1024)} MB", Color.Gray);
            Log($"Filtrando apenas arquivos dentro de: {bdocPath}", Color.Gray);

            allBdfFiles = await Task.Run(() =>
            {
                // Passar bdocPath para filtrar apenas arquivos dentro do diretório BDOC
                // Isso reduz drasticamente o uso de memória
                return MftScanner.ScanVolume(driveLetter, ".BDF", bdocPath, count =>
                {
                    if (count > 0)
                    {
                        UpdateProgress(count / 1000 % 100, 100, $"Fase 1: Lendo MFT ({count:N0} registros)");
                    }
                    else if (count < 0)
                    {
                        // Negativo indica resolução de paths
                        UpdateProgress(50, 100, $"Fase 1: Resolvendo caminhos ({-count:N0} arquivos)");
                    }
                }, memoryLimit);
            }, ct);
        }
        catch (InvalidOperationException ex)
        {
            Log($"ERRO: {ex.Message}", Color.Red);
            Log("Alternativa: Use o método '.NET EnumerateFiles'.", Color.Yellow);
            return;
        }
        catch (OutOfMemoryException)
        {
            Log("ERRO: Memória insuficiente. Tente o método '.NET EnumerateFiles'.", Color.Red);
            return;
        }

        Log($"Total de arquivos .BDF no volume: {allBdfFiles.Count:N0}", Color.Green);

        // Filtrar apenas arquivos dentro dos systemPaths
        var normalizedSystemPaths = systemPaths.Select(p => p.ToUpperInvariant().TrimEnd('\\')).ToList();

        var relevantFiles = allBdfFiles
            .Where(f =>
            {
                var upperPath = f.FullPath.ToUpperInvariant();
                return normalizedSystemPaths.Any(sp => upperPath.StartsWith(sp + "\\"));
            })
            .ToList();

        Log($"Arquivos .BDF dentro do sistema: {relevantFiles.Count:N0}", Color.Green);

        // Agrupar por tabela (extrair nome da tabela do caminho)
        var filesByTable = new ConcurrentDictionary<string, ConcurrentBag<FileHandleInfo>>();

        foreach (var file in relevantFiles)
        {
            ct.ThrowIfCancellationRequested();

            // Extrair tabela do caminho: ...\SISTEMA\TABELA\...
            foreach (var sp in normalizedSystemPaths)
            {
                var upperPath = file.FullPath.ToUpperInvariant();
                if (upperPath.StartsWith(sp + "\\"))
                {
                    var relativePath = file.FullPath.Substring(sp.Length + 1);
                    var tableName = relativePath.Split('\\')[0];

                    if (TryExtractHandle(file.FullPath, out var handle, out var fileName))
                    {
                        var bag = filesByTable.GetOrAdd(tableName, _ => new ConcurrentBag<FileHandleInfo>());
                        bag.Add(new FileHandleInfo(file.FullPath, fileName, handle));
                    }
                    break;
                }
            }
        }

        _phase1Duration = DateTime.Now - phase1Start;
        _totalFilesAnalyzed = relevantFiles.Count;

        Log($"Tabelas encontradas: {filesByTable.Count}", Color.Green);
        Log($"Fase 1 concluída em: {_phase1Duration:mm\\:ss\\.fff}", Color.Green);

        // ================================================================
        // FASE 2: Validação paralela no banco de dados
        // ================================================================
        Log("");
        Log("=== FASE 2: Validação paralela no banco de dados ===", Color.Cyan);
        var phase2Start = DateTime.Now;

        var tablesToProcess = filesByTable.Where(kvp => !kvp.Value.IsEmpty).ToList();
        var tablesProcessed = 0;
        var totalTables = tablesToProcess.Count;

        Log($"Tabelas com arquivos para validar: {totalTables}");

        await Parallel.ForEachAsync(tablesToProcess,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallelism, CancellationToken = ct },
            async (kvp, token) =>
            {
                var tableName = kvp.Key;
                var files = kvp.Value.ToList();
                var handleList = files.Select(f => f.Handle).Distinct().ToList();

                try
                {
                    var existingHandles = await QueryExistingHandlesAsync(
                        connectionString, tableName, handleList, token);

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
                        Log($"  [{tableName}] {files.Count} arquivos, {orphansInTable} órfãos", Color.Red);
                    }
                }
                catch (DbException ex)
                {
                    Log($"  [AVISO] Erro na tabela {tableName}: {ex.Message}", Color.Yellow);
                }

                var processed = Interlocked.Increment(ref tablesProcessed);
                Interlocked.Increment(ref _totalTablesProcessed);
                UpdateProgress(processed, totalTables, $"Fase 2: Validando ({processed}/{totalTables} tabelas)");
            });

        _phase2Duration = DateTime.Now - phase2Start;
        Log($"Fase 2 concluída em: {_phase2Duration:mm\\:ss\\.fff}", Color.Green);
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
        Log("=== FASE 1: Enumeração paralela de arquivos ===", Color.Cyan);
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
        Log($"Fase 1 concluída em: {_phase1Duration:mm\\:ss\\.fff}", Color.Green);

        // ================================================================
        // FASE 2: Validação paralela no banco de dados
        // ================================================================
        Log("");
        Log("=== FASE 2: Validação paralela no banco de dados ===", Color.Cyan);
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
                        Log($"  [{tableName}] {files.Count} arquivos, {orphansInTable} órfãos", Color.Red);
                    }
                }
                catch (DbException ex)
                {
                    Log($"  [AVISO] Erro na tabela {tableName}: {ex.Message}", Color.Yellow);
                }

                var processed = Interlocked.Increment(ref tablesProcessed);
                Interlocked.Increment(ref _totalTablesProcessed);
                UpdateProgress(processed, totalTables, $"Fase 2: Validando ({processed}/{totalTables} tabelas)");
            });

        _phase2Duration = DateTime.Now - phase2Start;
        Log($"Fase 2 concluída em: {_phase2Duration:mm\\:ss\\.fff}", Color.Green);
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
    /// </summary>
    private async Task<HashSet<int>> QueryExistingHandlesAsync(
        string connectionString,
        string tableName,
        List<int> handles,
        CancellationToken ct)
    {
        var existingHandles = new HashSet<int>();
        var batchSize = GetBatchSize();

        using var connection = CreateConnection(connectionString);
        await connection.OpenAsync(ct);

        for (int i = 0; i < handles.Count; i += batchSize)
        {
            ct.ThrowIfCancellationRequested();

            var batch = handles.Skip(i).Take(batchSize).ToList();
            var inClause = string.Join(",", batch);

            var query = BuildHandleQuery(tableName, inClause);

            using var cmd = connection.CreateCommand();
            cmd.CommandText = query;
            cmd.CommandTimeout = DB_COMMAND_TIMEOUT;

            using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                if (!reader.IsDBNull(0))
                    existingHandles.Add(reader.GetInt32(0));
            }
        }

        return existingHandles;
    }

    private void ExportResults(string filePath)
    {
        try
        {
            var sb = new StringBuilder();

            // Log completo do processamento
            sb.AppendLine("============================================");
            sb.AppendLine("              LOG DE EXECUÇÃO");
            sb.AppendLine("============================================");
            sb.AppendLine(txtLog.Text);
            sb.AppendLine();

            // Lista de arquivos órfãos
            var orphanList = _orphanFiles.OrderBy(o => o.Tabela).ThenBy(o => o.Handle).ToList();
            sb.AppendLine("============================================");
            sb.AppendLine("     LISTA DE ARQUIVOS ÓRFÃOS");
            sb.AppendLine("============================================");
            sb.AppendLine($"Total: {orphanList.Count} arquivos | Espaço: {FormatSize(_totalOrphanSize)}");
            sb.AppendLine();

            foreach (var orphan in orphanList)
            {
                sb.AppendLine($"Arquivo: {orphan.FileName} | Caminho: {orphan.FilePath} | Sistema: {orphan.Sistema} | Tabela: {orphan.Tabela} | Handle: {orphan.Handle} | Tamanho: {FormatSize(orphan.Size)}");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            Log($"Relatório exportado: {filePath}", Color.Green);
        }
        catch (Exception ex)
        {
            Log($"Erro ao exportar relatório: {ex.Message}", Color.Red);
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
                txtAppServer.Text,
                cmbSistema.Text,
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
                // Novo formato (6 linhas): bdoc, appServer, sistema, metodo, threads, backup
                if (lines.Length >= 3)
                {
                    txtBdocPath.Text = lines[0];
                    txtAppServer.Text = lines[1];
                    // Sistema: adicionar como item e selecionar (será recarregado ao conectar)
                    if (!string.IsNullOrWhiteSpace(lines[2]))
                    {
                        cmbSistema.Items.Add(lines[2]);
                        cmbSistema.SelectedIndex = 0;
                    }
                }
                if (lines.Length >= 4 && int.TryParse(lines[3], out var index))
                {
                    cmbSearchMethod.SelectedIndex = Math.Min(index, cmbSearchMethod.Items.Count - 1);
                }
                if (lines.Length >= 5 && decimal.TryParse(lines[4], out var parallelism))
                {
                    numParallelism.Value = Math.Clamp(parallelism, numParallelism.Minimum, numParallelism.Maximum);
                }
                if (lines.Length >= 6)
                {
                    txtBackupPath.Text = lines[5];
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
