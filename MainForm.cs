using System.Collections.Concurrent;
using System.Data.Common;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Text;
using Oracle.ManagedDataAccess.Client;

namespace BDFOrphanFinder;

#region Win32 API Interop
internal static class NativeMethods
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WIN32_FIND_DATA
    {
        public uint dwFileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternateFileName;
    }

    public const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
    public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindFirstFileW(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool FindNextFileW(IntPtr hFindFile, out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool FindClose(IntPtr hFindFile);

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
    /// <param name="basePath">Caminho base para filtrar. Se fornecido, apenas arquivos dentro deste path serão retornados.</param>
    /// <param name="progressCallback">Callback de progresso</param>
    /// <param name="memoryLimitBytes">Limite de memória em bytes</param>
    public static List<MftFileInfo> ScanVolume(
        char driveLetter,
        string extension,
        string? basePath = null,
        Action<int>? progressCallback = null,
        long memoryLimitBytes = DEFAULT_MEMORY_LIMIT_BYTES)
    {
        var files = new List<MftFileInfo>();

        // Usar dicionário com capacidade inicial menor para economizar memória
        // Se temos basePath, precisamos de muito menos entradas
        var initialCapacity = string.IsNullOrEmpty(basePath) ? 500000 : 50000;
        var directories = new Dictionary<long, DirectoryEntry>(initialCapacity);

        // Preparar filtro de basePath
        string[]? basePathParts = null;
        if (!string.IsNullOrEmpty(basePath))
        {
            // Normalizar basePath: remover drive letter e separar em partes
            // Ex: "B:\BDOC\sistema" -> ["BDOC", "sistema"]
            var pathWithoutDrive = basePath.Length > 2 && basePath[1] == ':'
                ? basePath.Substring(3)  // Remove "B:\"
                : basePath;
            basePathParts = pathWithoutDrive.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
        }

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
            var error = Marshal.GetLastWin32Error();
            throw new InvalidOperationException($"Nao foi possivel acessar o volume {driveLetter}:. Erro: {error}. Execute como Administrador.");
        }

        try
        {
            var mftData = new NativeMethods.MFT_ENUM_DATA_V0
            {
                StartFileReferenceNumber = 0,
                LowUsn = 0,
                HighUsn = long.MaxValue
            };

            const int bufferSize = 64 * 1024; // 64KB buffer (reduzido)
            var buffer = Marshal.AllocHGlobal(bufferSize);
            var extUpper = extension.ToUpperInvariant().TrimStart('*');

            // Cache de diretórios que estão no caminho do basePath (para otimização)
            // Key: fileRef, Value: true se está no caminho ou é ancestral do basePath
            HashSet<long>? relevantDirs = basePathParts != null ? new HashSet<long>() : null;

            try
            {
                int recordCount = 0;
                long estimatedMemoryUsage = 0;
                bool memoryWarningLogged = false;

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

                    // Verificar limite de memória periodicamente
                    if (recordCount % 50000 == 0)
                    {
                        estimatedMemoryUsage = (long)directories.Count * ESTIMATED_DIR_ENTRY_SIZE +
                                              (long)files.Count * ESTIMATED_FILE_ENTRY_SIZE;

                        if (estimatedMemoryUsage > memoryLimitBytes)
                        {
                            if (!memoryWarningLogged)
                            {
                                memoryWarningLogged = true;
                            }

                            // Limpar diretórios que provavelmente não serão usados
                            if (directories.Count > 1000000)
                            {
                                CompactDirectories(directories, files);
                            }
                        }
                    }

                    mftData.StartFileReferenceNumber = Marshal.ReadInt64(buffer, 0);
                    int offset = 8;

                    while (offset < bytesReturned)
                    {
                        var recordLength = Marshal.ReadInt32(buffer, offset);
                        if (recordLength == 0) break;

                        var fileRef = Marshal.ReadInt64(buffer, offset + 8);
                        var parentRef = Marshal.ReadInt64(buffer, offset + 16);
                        var fileAttributes = (uint)Marshal.ReadInt32(buffer, offset + 52);
                        var fileNameLength = Marshal.ReadInt16(buffer, offset + 56);
                        var fileNameOffset = Marshal.ReadInt16(buffer, offset + 58);

                        var fileName = Marshal.PtrToStringUni(buffer + offset + fileNameOffset, fileNameLength / 2);

                        if (!string.IsNullOrEmpty(fileName) && fileName != "." && fileName != "..")
                        {
                            var isDirectory = (fileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) != 0;

                            if (isDirectory)
                            {
                                // Se temos basePath, só armazenar diretórios que podem ser relevantes
                                if (basePathParts != null)
                                {
                                    // Armazenar sempre - vamos filtrar depois na resolução de paths
                                    // Mas com capacidade reduzida já economiza memória
                                    directories[fileRef] = new DirectoryEntry(fileName, parentRef);
                                }
                                else
                                {
                                    directories[fileRef] = new DirectoryEntry(fileName, parentRef);
                                }
                            }
                            else if (fileName.EndsWith(extUpper, StringComparison.OrdinalIgnoreCase))
                            {
                                files.Add(new MftFileInfo
                                {
                                    FileReferenceNumber = fileRef,
                                    ParentFileReferenceNumber = parentRef,
                                    FileName = fileName
                                });
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

        // Forçar GC para liberar memória
        GC.Collect(2, GCCollectionMode.Optimized);

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
        GC.Collect(1, GCCollectionMode.Optimized);
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
    private TextBox txtSystemName = null!;
    private ComboBox cmbDatabaseType = null!;
    private TextBox txtServer = null!;
    private TextBox txtDatabase = null!;
    private TextBox txtPort = null!;
    private TextBox txtUsername = null!;
    private TextBox txtPassword = null!;
    private ComboBox cmbSearchMethod = null!;
    private NumericUpDown numParallelism = null!;
    private TextBox txtBackupPath = null!;
    private CheckBox chkRemoveAfterBackup = null!;
    private Button btnBrowse = null!;
    private Button btnBrowseBackup = null!;
    private Button btnStart = null!;
    private Button btnCancel = null!;
    private ProgressBar progressBar = null!;
    private Label lblStatus = null!;
    private Label lblProgress = null!;
    private RichTextBox txtLog = null!;

    // Labels que mudam dinamicamente conforme o tipo de banco
    private Label lblServer = null!;
    private Label lblDatabase = null!;
    private Label lblPort = null!;

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

    private bool IsOracle => cmbDatabaseType.SelectedIndex == 1;

    public MainForm()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Identificador de Arquivos BDF Orfaos";
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
            RowCount = 8,
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

        // Row 1: Tipo Banco + Servidor
        configPanel.Controls.Add(new Label { Text = "Tipo Banco:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 1);
        cmbDatabaseType = new ComboBox
        {
            Dock = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Standard,
            BackColor = Color.White,
            Font = textBoxFont
        };
        cmbDatabaseType.Items.AddRange(new object[] { "SQL Server", "Oracle" });
        cmbDatabaseType.SelectedIndex = 0;
        cmbDatabaseType.SelectedIndexChanged += CmbDatabaseType_SelectedIndexChanged;
        configPanel.Controls.Add(cmbDatabaseType, 1, 1);
        lblServer = new Label { Text = "Servidor SQL:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont };
        configPanel.Controls.Add(lblServer, 3, 1);
        txtServer = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtServer, 4, 1);

        // Row 2: Database/Service Name + Port
        lblDatabase = new Label { Text = "Banco de Dados:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont };
        configPanel.Controls.Add(lblDatabase, 0, 2);
        txtDatabase = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtDatabase, 1, 2);
        lblPort = new Label { Text = "Porta:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont, Visible = false };
        configPanel.Controls.Add(lblPort, 3, 2);
        txtPort = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont, Text = "1521", Visible = false };
        configPanel.Controls.Add(txtPort, 4, 2);

        // Row 3: Username + Password
        configPanel.Controls.Add(new Label { Text = "Usuario:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 3);
        txtUsername = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtUsername, 1, 3);
        configPanel.Controls.Add(new Label { Text = "Senha:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 3);
        txtPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtPassword, 4, 3);

        // Row 4: Search Method + Parallelism
        configPanel.Controls.Add(new Label { Text = "Metodo de Busca:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 4);
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
        configPanel.Controls.Add(cmbSearchMethod, 1, 4);

        configPanel.Controls.Add(new Label { Text = "Threads:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginRight, ForeColor = labelColor, Font = labelFont }, 3, 4);
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
        configPanel.Controls.Add(numParallelism, 4, 4);

        // Row 5: Backup Path
        configPanel.Controls.Add(new Label { Text = "Caminho Backup BDF:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = labelMarginLeft, ForeColor = labelColor, Font = labelFont }, 0, 5);
        txtBackupPath = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Font = textBoxFont };
        configPanel.Controls.Add(txtBackupPath, 1, 5);
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
        configPanel.Controls.Add(btnBrowseBackup, 2, 5);

        // Checkbox para remover após backup (na coluna 3, ao lado do botão ...)
        chkRemoveAfterBackup = new CheckBox
        {
            Text = "Remover orfaos",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(10, 6, 3, 0),
            ForeColor = Color.FromArgb(180, 0, 0),
            Font = labelFont,
            Cursor = Cursors.Hand
        };
        configPanel.Controls.Add(chkRemoveAfterBackup, 3, 5);
        configPanel.SetColumnSpan(chkRemoveAfterBackup, 2);

        // Row 6: Buttons
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

        configPanel.Controls.Add(buttonPanel, 1, 6);
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

    private void CmbDatabaseType_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (IsOracle)
        {
            lblServer.Text = "Servidor Oracle:";
            lblDatabase.Text = "Service Name:";
            lblPort.Visible = true;
            txtPort.Visible = true;
        }
        else
        {
            lblServer.Text = "Servidor SQL:";
            lblDatabase.Text = "Banco de Dados:";
            lblPort.Visible = false;
            txtPort.Visible = false;
        }
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

        var dbType = IsOracle ? "Oracle" : "SQL Server";

        if (string.IsNullOrWhiteSpace(txtServer.Text) ||
            string.IsNullOrWhiteSpace(txtDatabase.Text) ||
            string.IsNullOrWhiteSpace(txtUsername.Text))
        {
            MessageBox.Show($"Dados de conexao {dbType} incompletos.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (IsOracle && string.IsNullOrWhiteSpace(txtPort.Text))
        {
            MessageBox.Show("Porta do Oracle e obrigatoria.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (!string.IsNullOrWhiteSpace(txtBackupPath.Text) && !Directory.Exists(txtBackupPath.Text))
        {
            MessageBox.Show("Diretorio de backup invalido ou nao existe.", "Validacao",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        // Se checkbox de remover após backup está marcado, o caminho de backup é obrigatório
        if (chkRemoveAfterBackup.Checked && string.IsNullOrWhiteSpace(txtBackupPath.Text))
        {
            MessageBox.Show("Para remover orfaos apos backup, e necessario configurar um caminho de backup valido.", "Validacao",
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
        cmbDatabaseType.Enabled = !running;
        txtServer.Enabled = !running;
        txtDatabase.Enabled = !running;
        txtPort.Enabled = !running;
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

    // ===== Metodos auxiliares de abstração de banco =====

    private string BuildConnectionString(int maxParallelism)
    {
        if (IsOracle)
        {
            return $"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST={txtServer.Text})(PORT={txtPort.Text})))(CONNECT_DATA=(SERVICE_NAME={txtDatabase.Text})));User Id={txtUsername.Text};Password={txtPassword.Text};Max Pool Size={maxParallelism + 5};";
        }
        else
        {
            return $"Server={txtServer.Text};Database={txtDatabase.Text};User Id={txtUsername.Text};Password={txtPassword.Text};TrustServerCertificate=True;Max Pool Size={maxParallelism + 5};";
        }
    }

    private DbConnection CreateConnection(string connectionString)
    {
        if (IsOracle)
            return new OracleConnection(connectionString);
        return new SqlConnection(connectionString);
    }

    private string BuildHandleQuery(string tableName, string? inClause = null)
    {
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
        // Oracle tem limite de 1000 expressoes na clausula IN
        // SQL Server suporta ate ~2000
        return IsOracle ? 1000 : 2000;
    }

    private async Task ProcessOrphanFilesAsync(CancellationToken ct)
    {
        var bdocPath = txtBdocPath.Text;
        var systemName = txtSystemName.Text;
        var maxParallelism = (int)numParallelism.Value;
        var searchMethod = cmbSearchMethod.SelectedIndex;
        var dbType = IsOracle ? "Oracle" : "SQL Server";

        Log("Iniciando processamento...", Color.Cyan);
        Log($"Diretorio BDOC: {bdocPath}");
        Log($"Sistema: {systemName}");
        Log($"Banco de dados: {dbType}");
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

        // Connection string
        var connectionString = BuildConnectionString(maxParallelism);

        // Test connection
        Log($"Testando conexao com banco de dados {dbType}...");
        using (var testConn = CreateConnection(connectionString))
        {
            await testConn.OpenAsync(ct);
            Log("Conexao estabelecida com sucesso!", Color.Green);
        }

        // Use strategy based on selection
        // 0 = MFT (Master File Table) - requer Admin
        // 1 = .NET EnumerateFiles
        if (searchMethod == 0) // MFT
        {
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

        if (!string.IsNullOrWhiteSpace(backupPath) && !_orphanFiles.IsEmpty)
        {
            // Se checkbox de remoção está marcado, pedir confirmação
            if (chkRemoveAfterBackup.Checked)
            {
                var confirmResult = MessageBox.Show(
                    $"Foram encontrados {_orphanFiles.Count} arquivos orfaos ({FormatSize(_totalOrphanSize)}).\n\n" +
                    $"Os arquivos serao copiados para:\n{backupPath}\n\n" +
                    "ATENCAO: Apos o backup, os arquivos originais serao REMOVIDOS do BDOC!\n\n" +
                    "Deseja continuar com o backup e remocao?",
                    "Confirmacao de Remocao",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (confirmResult != DialogResult.Yes)
                {
                    Log("Operacao de backup/remocao cancelada pelo usuario.", Color.Yellow);
                    // Pular para o relatório final
                    goto FinalReport;
                }

                shouldRemoveOrphans = true;
            }

            Log("");
            Log("=== FASE 3: Backup dos arquivos orfaos ===", Color.Cyan);
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
                        var relativePath = Path.GetRelativePath(bdocPath, orphan.FilePath);
                        var destPath = Path.Combine(backupPath, relativePath);
                        var destDir = Path.GetDirectoryName(destPath)!;

                        Directory.CreateDirectory(destDir);
                        File.Copy(orphan.FilePath, destPath, overwrite: false);

                        // Marcar como copiado com sucesso
                        successfullyCopied.Add(orphan.FilePath);

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

            // ================================================================
            // FASE 4: Remoção dos arquivos órfãos do BDOC (se solicitado)
            // ================================================================
            if (shouldRemoveOrphans && !successfullyCopied.IsEmpty)
            {
                Log("");
                Log("=== FASE 4: Remocao dos arquivos orfaos do BDOC ===", Color.Cyan);
                var phase4Start = DateTime.Now;

                var filesToRemove = successfullyCopied.ToList();
                var totalToRemove = filesToRemove.Count;
                var dirsRemoved = 0;

                // Normalizar systemPaths para comparação (limite de deleção de pastas)
                var normalizedSystemPaths = systemPaths.Select(p => p.TrimEnd('\\').ToUpperInvariant()).ToHashSet();

                Log($"Removendo {totalToRemove} arquivos do BDOC...");

                await Parallel.ForEachAsync(filesToRemove,
                    new ParallelOptions { MaxDegreeOfParallelism = (int)numParallelism.Value, CancellationToken = ct },
                    async (filePath, token) =>
                    {
                        try
                        {
                            // Deletar o arquivo
                            File.Delete(filePath);
                            var done = Interlocked.Increment(ref filesRemoved);

                            // Tentar remover pastas vazias subindo até o diretório da tabela
                            var currentDir = Path.GetDirectoryName(filePath);
                            while (!string.IsNullOrEmpty(currentDir))
                            {
                                var normalizedCurrentDir = currentDir.TrimEnd('\\').ToUpperInvariant();

                                // Parar se chegou no diretório do sistema (ex: B:\BDOC\2025-12\RH_PROD)
                                // Não deletar a pasta da tabela nem acima
                                var isSystemPath = normalizedSystemPaths.Any(sp =>
                                    normalizedCurrentDir.Equals(sp) ||
                                    normalizedCurrentDir.StartsWith(sp + "\\") &&
                                    normalizedCurrentDir.Substring(sp.Length + 1).IndexOf('\\') == -1);

                                if (isSystemPath || normalizedSystemPaths.Any(sp => sp.StartsWith(normalizedCurrentDir)))
                                    break;

                                try
                                {
                                    // Só deleta se estiver vazio
                                    if (Directory.Exists(currentDir) && !Directory.EnumerateFileSystemEntries(currentDir).Any())
                                    {
                                        Directory.Delete(currentDir);
                                        Interlocked.Increment(ref dirsRemoved);
                                        currentDir = Path.GetDirectoryName(currentDir);
                                    }
                                    else
                                    {
                                        break; // Pasta não está vazia, parar
                                    }
                                }
                                catch
                                {
                                    break; // Erro ao deletar pasta, parar
                                }
                            }

                            if (done % 100 == 0 || done == totalToRemove)
                            {
                                UpdateProgress(done, totalToRemove, $"Fase 4: Removendo ({done}/{totalToRemove} arquivos)");
                            }
                        }
                        catch (Exception ex)
                        {
                            Interlocked.Increment(ref removeErrors);
                            Log($"  [ERRO REMOCAO] {Path.GetFileName(filePath)}: {ex.Message}", Color.Yellow);
                        }

                        await Task.CompletedTask;
                    });

                var phase4Duration = DateTime.Now - phase4Start;
                Log($"Remocao concluida: {filesRemoved} arquivos, {dirsRemoved} pastas vazias, {removeErrors} erros", filesRemoved > 0 ? Color.Green : Color.Yellow);
                Log($"Fase 4 concluida em: {phase4Duration:mm\\:ss\\.fff}", Color.Green);
            }
        }

        FinalReport:

        // Final report
        var elapsed = DateTime.Now - startTime;
        var orphanCount = _orphanFiles.Count;

        Log("", Color.White);
        Log("============================================", Color.Cyan);
        Log("           RELATORIO FINAL", Color.Cyan);
        Log("============================================", Color.Cyan);
        Log($"Banco de dados: {dbType}");
        Log($"Total de arquivos analisados: {_totalFilesAnalyzed}");
        Log($"Arquivos COM registro na base: {_filesWithRecord}", Color.Green);
        Log($"Arquivos ORFAOS (sem registro): {orphanCount}", orphanCount > 0 ? Color.Red : Color.Green);
        Log($"Espaco ocupado por orfaos: {FormatSize(_totalOrphanSize)}", Color.Yellow);
        if (_totalFilesCopied > 0 || _totalCopyErrors > 0)
        {
            Log($"Backup: {_totalFilesCopied} copiados, {_totalCopyErrors} erros", Color.Cyan);
        }
        if (filesRemoved > 0 || removeErrors > 0)
        {
            Log($"Remocao: {filesRemoved} removidos, {removeErrors} erros", Color.Cyan);
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
        Log("NOTA: Este metodo requer execucao como Administrador", Color.Yellow);
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
            Log($"Limite de memoria configurado: {memoryLimit / (1024 * 1024)} MB", Color.Gray);
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
            Log("Alternativa: Use o metodo 'Win32 API' que nao requer Admin.", Color.Yellow);
            return;
        }
        catch (OutOfMemoryException)
        {
            Log("ERRO: Memoria insuficiente. Tente o metodo 'Win32 API'.", Color.Red);
            GC.Collect(2, GCCollectionMode.Forced);
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
                        Log($"  [{tableName}] {files.Count} arquivos, {orphansInTable} orfaos", Color.Red);
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
        Log($"Fase 2 concluida em: {_phase2Duration:mm\\:ss\\.fff}", Color.Green);
    }

    /// <summary>
    /// ESTRATÉGIA WIN32 API - MÁXIMA PERFORMANCE
    /// Usa FindFirstFile/FindNextFile para enumeração mais rápida
    /// </summary>
    private async Task ProcessWithWin32StrategyAsync(
        List<string> systemPaths,
        string systemName,
        string connectionString,
        int maxParallelism,
        CancellationToken ct)
    {
        // ================================================================
        // FASE 1: Enumeração com Win32 API por diretório de tabela
        // ================================================================
        Log("");
        Log("=== FASE 1: Enumeracao com Win32 API (Alta Performance) ===", Color.Cyan);
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

        // Processar diretórios em paralelo usando Win32 API
        var allDirs = tableDirs.SelectMany(kvp => kvp.Value.Select(dir => (Table: kvp.Key, Dir: dir))).ToList();

        await Parallel.ForEachAsync(allDirs,
            new ParallelOptions { MaxDegreeOfParallelism = maxParallelism, CancellationToken = ct },
            async (item, token) =>
            {
                var (tableName, tableDir) = item;
                token.ThrowIfCancellationRequested();

                var bag = filesByTable.GetOrAdd(tableName, _ => new ConcurrentBag<FileHandleInfo>());
                var localFileCount = 0;

                try
                {
                    // Usar Win32 API para enumerar arquivos
                    EnumerateFilesWithWin32Api(tableDir, "*.BDF", bag, ref localFileCount);
                }
                catch (Exception) { }

                Interlocked.Add(ref totalFilesFound, localFileCount);
                var processed = Interlocked.Increment(ref directoriesProcessed);

                if (processed % 10 == 0 || processed == totalDirectories)
                {
                    UpdateProgress(processed, totalDirectories, $"Fase 1: Win32 API ({processed}/{totalDirectories} dirs, {totalFilesFound} arquivos)");
                }

                await Task.CompletedTask;
            });

        _phase1Duration = DateTime.Now - phase1Start;
        _totalFilesAnalyzed = totalFilesFound;

        Log($"Arquivos encontrados: {totalFilesFound} em {tableDirs.Count} tabelas", Color.Green);
        Log($"Fase 1 concluida em: {_phase1Duration:mm\\:ss\\.fff}", Color.Green);

        // ================================================================
        // FASE 2: Validação paralela no banco de dados (igual ao híbrido)
        // ================================================================
        Log("");
        Log("=== FASE 2: Validacao paralela no banco de dados ===", Color.Cyan);
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
                        Log($"  [{tableName}] {files.Count} arquivos, {orphansInTable} orfaos", Color.Red);
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
        Log($"Fase 2 concluida em: {_phase2Duration:mm\\:ss\\.fff}", Color.Green);
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
                catch (DbException ex)
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

        using var connection = CreateConnection(connectionString);
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
                var query = BuildHandleQuery(table);
                using var cmd = connection.CreateCommand();
                cmd.CommandText = query;
                cmd.CommandTimeout = 300;

                using var reader = await cmd.ExecuteReaderAsync(ct);
                while (await reader.ReadAsync(ct))
                {
                    if (!reader.IsDBNull(0))
                        handles.Add(reader.GetInt32(0));
                }
            }
            catch (DbException ex)
            {
                Log($"  [AVISO] Erro ao consultar tabela {table}: {ex.Message}", Color.Yellow);
                continue;
            }

            if (handles.Count == 0) continue;

            foreach (var tablePath in tablePaths[table])
            {
                ct.ThrowIfCancellationRequested();
                if (!Directory.Exists(tablePath)) continue;

                // searchMethod: 3 = .NET EnumerateFiles, 4 = Robocopy
                var files = searchMethod == 3 ? GetFilesWithDotNet(tablePath) : GetFilesWithRobocopy(tablePath);

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
        // Método auxiliar - usa .NET EnumerateFiles como padrão
        return GetFilesWithDotNet(path);
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

    /// <summary>
    /// Enumeracao de arquivos usando Win32 API (FindFirstFile/FindNextFile)
    /// Mais rapido que Directory.EnumerateFiles para grandes volumes de arquivos
    /// </summary>
    private static IEnumerable<string> GetFilesWithWin32Api(string path, string pattern = "*.BDF")
    {
        var files = new List<string>();
        var directoriesToProcess = new Stack<string>();
        directoriesToProcess.Push(path);

        while (directoriesToProcess.Count > 0)
        {
            var currentDir = directoriesToProcess.Pop();

            // Primeiro, buscar subdiretorios
            var searchPathDirs = Path.Combine(currentDir, "*");
            var findHandleDirs = NativeMethods.FindFirstFileW(searchPathDirs, out var findDataDirs);

            if (findHandleDirs != NativeMethods.INVALID_HANDLE_VALUE)
            {
                try
                {
                    do
                    {
                        if (findDataDirs.cFileName == "." || findDataDirs.cFileName == "..")
                            continue;

                        if ((findDataDirs.dwFileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) != 0)
                        {
                            directoriesToProcess.Push(Path.Combine(currentDir, findDataDirs.cFileName));
                        }
                    } while (NativeMethods.FindNextFileW(findHandleDirs, out findDataDirs));
                }
                finally
                {
                    NativeMethods.FindClose(findHandleDirs);
                }
            }

            // Depois, buscar arquivos com o pattern
            var searchPathFiles = Path.Combine(currentDir, pattern);
            var findHandleFiles = NativeMethods.FindFirstFileW(searchPathFiles, out var findDataFiles);

            if (findHandleFiles != NativeMethods.INVALID_HANDLE_VALUE)
            {
                try
                {
                    do
                    {
                        if ((findDataFiles.dwFileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) == 0)
                        {
                            files.Add(Path.Combine(currentDir, findDataFiles.cFileName));
                        }
                    } while (NativeMethods.FindNextFileW(findHandleFiles, out findDataFiles));
                }
                finally
                {
                    NativeMethods.FindClose(findHandleFiles);
                }
            }
        }

        return files;
    }

    /// <summary>
    /// Versao thread-safe que adiciona diretamente a uma ConcurrentBag
    /// </summary>
    private static void EnumerateFilesWithWin32Api(string path, string pattern, ConcurrentBag<FileHandleInfo> results, ref int fileCount)
    {
        var directoriesToProcess = new Stack<string>();
        directoriesToProcess.Push(path);

        while (directoriesToProcess.Count > 0)
        {
            var currentDir = directoriesToProcess.Pop();

            // Primeiro, buscar subdiretorios
            var searchPathDirs = Path.Combine(currentDir, "*");
            var findHandleDirs = NativeMethods.FindFirstFileW(searchPathDirs, out var findDataDirs);

            if (findHandleDirs != NativeMethods.INVALID_HANDLE_VALUE)
            {
                try
                {
                    do
                    {
                        if (findDataDirs.cFileName == "." || findDataDirs.cFileName == "..")
                            continue;

                        if ((findDataDirs.dwFileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) != 0)
                        {
                            directoriesToProcess.Push(Path.Combine(currentDir, findDataDirs.cFileName));
                        }
                    } while (NativeMethods.FindNextFileW(findHandleDirs, out findDataDirs));
                }
                finally
                {
                    NativeMethods.FindClose(findHandleDirs);
                }
            }

            // Depois, buscar arquivos com o pattern
            var searchPathFiles = Path.Combine(currentDir, pattern);
            var findHandleFiles = NativeMethods.FindFirstFileW(searchPathFiles, out var findDataFiles);

            if (findHandleFiles != NativeMethods.INVALID_HANDLE_VALUE)
            {
                try
                {
                    do
                    {
                        if ((findDataFiles.dwFileAttributes & NativeMethods.FILE_ATTRIBUTE_DIRECTORY) == 0)
                        {
                            var filePath = Path.Combine(currentDir, findDataFiles.cFileName);
                            if (TryExtractHandle(filePath, out var handle, out var fileName))
                            {
                                results.Add(new FileHandleInfo(filePath, fileName, handle));
                                Interlocked.Increment(ref fileCount);
                            }
                        }
                    } while (NativeMethods.FindNextFileW(findHandleFiles, out findDataFiles));
                }
                finally
                {
                    NativeMethods.FindClose(findHandleFiles);
                }
            }
        }
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
                txtBackupPath.Text,
                cmbDatabaseType.SelectedIndex.ToString(),
                txtPort.Text
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
                // Novos campos: tipo de banco e porta
                if (lines.Length >= 9 && int.TryParse(lines[8], out var dbType))
                {
                    cmbDatabaseType.SelectedIndex = Math.Min(dbType, cmbDatabaseType.Items.Count - 1);
                }
                if (lines.Length >= 10)
                {
                    txtPort.Text = lines[9];
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
