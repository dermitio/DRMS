using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Security.Principal;
using WindowsInput;
using WindowsInput.Native;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using System.Drawing;
using System.Net.Http;

// Represents a message with associated text, a weight for rarity selection, and an optional image URL.
public class Message
{
    public string Text { get; set; }
    public int Weight { get; set; }
    public string ImageUrl { get; set; }

    public Message(string text, int weight, string imageUrl = "")
    {
        Text = text;
        Weight = weight;
        ImageUrl = imageUrl;
    }
}

/// <summary>
/// Defines the application's customizable settings.
/// </summary>
public class Settings
{
    // Color settings (stored as ARGB integers for JSON serialization)
    public int BackColorArgb { get; set; } = Color.FromArgb(43, 42, 51).ToArgb();
    public int ForeColorArgb { get; set; } = Color.FromArgb(255, 255, 255).ToArgb();
    public int ControlBackColorArgb { get; set; } = Color.FromArgb(66, 65, 77).ToArgb();
    public int ButtonHoverColorArgb { get; set; } = Color.FromArgb(126, 125, 130).ToArgb();
    public int ButtonActiveColorArgb { get; set; } = Color.FromArgb(32, 35, 39).ToArgb();
    public int DataGridViewHeaderBackColorArgb { get; set; } = Color.FromArgb(32, 34, 37).ToArgb();
    public int HighlightColorArgb { get; set; } = Color.FromArgb(79, 159, 236).ToArgb();

    // Hotkey settings (stored as string representation of Keys enum)
    public string SendHotkey { get; set; } = "F5";
    public string SaveHotkey { get; set; } = "S"; // Ctrl+S
    public string LoadHotkey { get; set; } = "L"; // Ctrl+L
    public string AddHotkey { get; set; } = "N"; // Ctrl+N
    public string DuplicateHotkey { get; set; } = "D"; // Ctrl+D
    public string DeleteHotkey { get; set; } = "Delete";
    public string EditHotkey { get; set; } = "Enter"; // For DataGridView double-click

    // New setting for allowing duplicate messages
    public bool AllowDuplicateMessages { get; set; } = false;

    // Constructor to set default values if needed
    public Settings() { }
}

/// <summary>
/// Manages loading and saving application settings.
/// </summary>
public static class SettingsManager
{
    private static readonly string SettingsFilePath = Path.Combine(Application.StartupPath, "settings.json");
    public static Settings CurrentSettings { get; private set; }

    static SettingsManager()
    {
        LoadSettings();
    }

    public static void LoadSettings()
    {
        if (File.Exists(SettingsFilePath))
        {
            try
            {
                string json = File.ReadAllText(SettingsFilePath);
                CurrentSettings = JsonConvert.DeserializeObject<Settings>(json);
                // Ensure default values for new properties in existing settings files
                if (CurrentSettings.SendHotkey == null) CurrentSettings.SendHotkey = "F5";
                if (CurrentSettings.SaveHotkey == null) CurrentSettings.SaveHotkey = "S";
                if (CurrentSettings.LoadHotkey == null) CurrentSettings.LoadHotkey = "L";
                if (CurrentSettings.AddHotkey == null) CurrentSettings.AddHotkey = "N";
                if (CurrentSettings.DuplicateHotkey == null) CurrentSettings.DuplicateHotkey = "D";
                if (CurrentSettings.DeleteHotkey == null) CurrentSettings.DeleteHotkey = "Delete";
                if (CurrentSettings.EditHotkey == null) CurrentSettings.EditHotkey = "Enter";
                // Ensure new property has a default value if not present in loaded file
                // Use a try-catch for GetType().GetProperty to avoid error if property doesn't exist in older JSON
                if (CurrentSettings.GetType().GetProperty(nameof(Settings.AllowDuplicateMessages)) == null)
                {
                    CurrentSettings.AllowDuplicateMessages = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading settings: {ex.Message}\nDefault settings will be used.", "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CurrentSettings = new Settings(); // Fallback to default
            }
        }
        else
        {
            CurrentSettings = new Settings(); // Create default settings if file doesn't exist
        }
    }

    public static void SaveSettings(Settings settings)
    {
        try
        {
            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving settings: {ex.Message}", "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // Helper methods to get colors from settings
    public static Color GetBackColor() => Color.FromArgb(CurrentSettings.BackColorArgb);
    public static Color GetForeColor() => Color.FromArgb(CurrentSettings.ForeColorArgb);
    public static Color GetControlBackColor() => Color.FromArgb(CurrentSettings.ControlBackColorArgb);
    public static Color GetButtonHoverColor() => Color.FromArgb(CurrentSettings.ButtonHoverColorArgb);
    public static Color GetButtonActiveColor() => Color.FromArgb(CurrentSettings.ButtonActiveColorArgb);
    public static Color GetDataGridViewHeaderBackColor() => Color.FromArgb(CurrentSettings.DataGridViewHeaderBackColorArgb);
    public static Color GetHighlightColor() => Color.FromArgb(CurrentSettings.HighlightColorArgb);
}

// Define the main form for the application
class MainForm : Form
{
    // DllImport attributes for interacting with Windows API functions
    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd); // Sets the window to the foreground

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); // Shows/restores a window

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId); // Gets the process ID associated with a window thread

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow(); // Gets the handle to the foreground window

    [DllImport("user32.dll")]
    static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach); // Attaches/detaches input processing of two threads

    [DllImport("kernel32.dll")]
    static extern uint GetCurrentThreadId(); // Gets the current thread ID

    // Constant for restoring a window to its normal size and position
    const int SW_RESTORE = 9;

    /// <summary>
    /// Forces a window to the foreground and makes it active.
    /// This uses native Windows API calls to ensure focus.
    /// </summary>
    /// <param name="hWnd">The handle to the window to focus.</param>
    static void ForceFocusWindow(IntPtr hWnd)
    {
        ShowWindow(hWnd, SW_RESTORE); // Restore the window if minimized or maximized

        // Get the thread ID of the foreground window and the current application
        uint foreThread = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
        uint appThread = GetCurrentThreadId();

        // Attach the current thread's input to the foreground window's thread.
        // This is necessary to gain foreground access reliably.
        AttachThreadInput(appThread, foreThread, true);
        SetForegroundWindow(hWnd); // Set the target window to foreground
        AttachThreadInput(appThread, foreThread, false); // Detach threads after focus is gained
    }

    /// <summary>
    /// Checks if the current application is running with administrative privileges.
    /// </summary>
    /// <returns>True if running as administrator, false otherwise.</returns>
    static bool IsRunningAsAdmin()
    {
        using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
        {
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }

    /// <summary>
    /// Relaunches the current application with administrative privileges.
    /// </summary>
    static void RelaunchAsAdmin()
    {
        var exeName = Process.GetCurrentProcess().MainModule.FileName; // Get the executable path
        var startInfo = new ProcessStartInfo(exeName)
        {
            UseShellExecute = true, // Must be true to use "runas" verb
            Verb = "runas"          // Specifies to run the process as administrator
        };

        try
        {
            Process.Start(startInfo); // Attempt to start the new process
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error relaunching as admin: {ex.Message}\nThis might mean the user canceled the User Account Control (UAC) prompt or another error occurred.", "Admin Relaunch Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        Environment.Exit(0); // Terminate the current non-admin process
    }

    /// <summary>
    /// Helper method to escape special characters for SendKeys.
    /// Characters like '{', '}', '(', ')', '[', ']', '+', '^', '%', '~' have special meaning.
    /// </summary>
    /// <param name="text">The original text to escape.</param>
    /// <returns>The escaped text suitable for SendKeys.</returns>
    static string EscapeSendKeysCharacters(string text)
    {
        return text.Replace("{", "{{}")
                   .Replace("}", "{}}")
                   .Replace("(", "{(}")
                   .Replace(")", "{)}")
                   .Replace("[", "{[}")
                   .Replace("]", "{]}")
                   .Replace("+", "{+}")
                   .Replace("^", "{^}")
                   .Replace("%", "{%}")
                   .Replace("~", "{~}");
    }

    // UI Controls for MainForm
    private DataGridView dgvMessages;
    private TextBox txtMessageText; // For adding new messages
    private NumericUpDown nudMessageWeight; // For adding new messages
    private TextBox txtImageUrl; // For adding image URLs
    private Button btnAdd;
    private Button btnUpdate;
    private Button btnDelete;
    private Button btnStartStop;
    private Label lblStatus;
    private Button btnSaveMessages; // New Save button for export
    private Button btnLoadMessages; // New Load button for import
    private NumericUpDown nudLoopCount; // New control for loop count
    private Label lblLoopCount; // Label for loop count
    private Label lblText;
    private Label lblWeight;
    private Label lblImageUrl;
    private RichTextBox rtbLog; // New: Debug log console
    private TextBox txtSearch; // New: Search/Filter messages
    private Label lblSearch; // New: Label for search
    private Button btnDuplicate; // New: Duplicate message button
    private Button btnSettings; // New: Settings button
    private ToolTip toolTip1;

    // Message lists
    private BindingSource messageBindingSource; // Used to bind messages to DataGridView
    private List<Message> allMessages;
    private List<Message> availableMessages;

    // Simulation related fields
    private InputSimulator inputSimulator;
    private Random random;
    private CancellationTokenSource cancellationTokenSource;
    private Task sendingTask;
    private bool isSending = false;

    // DataGridView Drag and Drop fields
    private Rectangle dragBoxFromMouseDown;
    private int rowIndexFromMouseDown;
    private int rowIndexOfItemUnderMouseToDrop;
    private DataGridViewRow _activeMessageRow; // Stores the currently sending row for highlighting

    // Constants for delays
    const int FocusDelayMs = 1000; // Delay after focusing Discord window
    const int InputDelayMs = 100;  // Delay after typing each line or sending command (reduced slightly)
    const int ImagePasteDelayMs = 750; // Delay specifically after pasting an image (increased)
    const int ClipboardClearDelayMs = 100; // Small delay before clearing clipboard after paste (increased)


    public MainForm()
    {
        InitializeComponent();
        ApplyTheme(); // Apply dark theme after initializing components
        InitializeData();
        Log("Application started."); // Log application start
    }

    // New: Log method to output messages to the RichTextBox
    private void Log(string message)
    {
        // Use Invoke to update UI from a non-UI thread safely
        if (rtbLog.InvokeRequired)
        {
            rtbLog.Invoke(new MethodInvoker(delegate
            {
                rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                rtbLog.ScrollToCaret(); // Auto-scroll to the latest message
            }));
        }
        else
        {
            rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            rtbLog.ScrollToCaret();
        }
    }


    private void InitializeComponent()
    {
        // Form Properties
        this.Text = "Discord Message Randomizer";
        this.Size = new System.Drawing.Size(800, 750); // Increased height to prevent cutoff
        this.MinimumSize = new System.Drawing.Size(700, 600); // Allow some resizing
        this.StartPosition = FormStartPosition.CenterScreen;
        this.KeyPreview = true; // Enable form to receive key events first
        this.KeyDown += MainForm_KeyDown; // Attach KeyDown event handler

        // ToolTip component initialization
        toolTip1 = new ToolTip();
        toolTip1.AutoPopDelay = 5000;
        toolTip1.InitialDelay = 1000;
        toolTip1.ReshowDelay = 500;
        toolTip1.ShowAlways = true;

        // Main layout panel for the whole form
        TableLayoutPanel mainLayout = new TableLayoutPanel();
        mainLayout.Dock = DockStyle.Fill; // Single column layout
        mainLayout.ColumnCount = 1;
        mainLayout.RowCount = 4; // Four rows: Search, DGV, Controls, Log
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Row 0: Search Panel height
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // Row 1: DataGridView fills remaining space
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 250F)); // Row 2: Control Panel height (increased to prevent status label cutoff)
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150F)); // Row 3: Log Console height
        mainLayout.BackColor = SettingsManager.GetBackColor(); // Set background for main layout

        // 1. Search Panel (for Row 0 of mainLayout)
        TableLayoutPanel searchPanel = new TableLayoutPanel();
        searchPanel.Dock = DockStyle.Fill; // Fill its cell in mainLayout
        searchPanel.Padding = new Padding(10, 5, 10, 5);
        searchPanel.ColumnCount = 2;
        searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80F));
        searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        searchPanel.BackColor = SettingsManager.GetControlBackColor();

        // Search Label
        lblSearch = new Label() { Text = "Search:", Anchor = AnchorStyles.Left };
        searchPanel.Controls.Add(lblSearch, 0, 0);

        // Search Textbox
        txtSearch = new TextBox() { Anchor = AnchorStyles.Left | AnchorStyles.Right, BorderStyle = BorderStyle.FixedSingle };
        txtSearch.TextChanged += TxtSearch_TextChanged;
        txtSearch.BackColor = SettingsManager.GetButtonActiveColor();
        txtSearch.ForeColor = SettingsManager.GetForeColor();
        toolTip1.SetToolTip(txtSearch, "Filter messages by text or image URL.");
        searchPanel.Controls.Add(txtSearch, 1, 0);
        mainLayout.Controls.Add(searchPanel, 0, 0); // Add search panel to Row 0


        // 2. DataGridView for messages (for Row 1 of mainLayout)
        dgvMessages = new DataGridView();
        dgvMessages.Dock = DockStyle.Fill; // Fill its cell in mainLayout
        dgvMessages.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dgvMessages.AllowUserToAddRows = false; // Prevent direct adding via DGV
        dgvMessages.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvMessages.MultiSelect = false;
        dgvMessages.ReadOnly = true; // Make it read-only for direct editing, use the separate edit window
        dgvMessages.SelectionChanged += DgvMessages_SelectionChanged;
        dgvMessages.CellDoubleClick += DgvMessages_CellDoubleClick; // For editing on double-click

        // Drag and drop event handlers
        dgvMessages.MouseDown += DgvMessages_MouseDown;
        dgvMessages.MouseMove += DgvMessages_MouseMove;
        dgvMessages.DragOver += DgvMessages_DragOver;
        dgvMessages.DragDrop += DgvMessages_DragDrop;
        dgvMessages.AllowDrop = true;

        // DataGridView styling for dark mode
        dgvMessages.BackgroundColor = SettingsManager.GetBackColor();
        dgvMessages.GridColor = SettingsManager.GetControlBackColor();
        dgvMessages.DefaultCellStyle.BackColor = SettingsManager.GetControlBackColor();
        dgvMessages.DefaultCellStyle.ForeColor = SettingsManager.GetForeColor();
        dgvMessages.DefaultCellStyle.SelectionBackColor = Color.FromArgb(70, 73, 79); // A subtle selection color
        dgvMessages.DefaultCellStyle.SelectionForeColor = SettingsManager.GetForeColor();
        dgvMessages.ColumnHeadersDefaultCellStyle.BackColor = SettingsManager.GetDataGridViewHeaderBackColor();
        dgvMessages.ColumnHeadersDefaultCellStyle.ForeColor = SettingsManager.GetForeColor();
        dgvMessages.EnableHeadersVisualStyles = false; // Required to apply custom header styles
        dgvMessages.RowHeadersDefaultCellStyle.BackColor = SettingsManager.GetDataGridViewHeaderBackColor();
        dgvMessages.RowHeadersDefaultCellStyle.ForeColor = SettingsManager.GetForeColor();
        mainLayout.Controls.Add(dgvMessages, 0, 1); // Add DGV to Row 1

        // DataGridView Column Customization
        // This must be done after setting the DataSource for auto-generated columns to exist.
        dgvMessages.DataBindingComplete += (sender, e) =>
        {
            if (dgvMessages.Columns.Contains("Text"))
            {
                dgvMessages.Columns["Text"].HeaderText = "Message"; // Changed to "Message"
            }
            if (dgvMessages.Columns.Contains("Weight"))
            {
                dgvMessages.Columns["Weight"].HeaderText = "Weight";
            }
            if (dgvMessages.Columns.Contains("ImageUrl"))
            {
                dgvMessages.Columns["ImageUrl"].HeaderText = "Image Path"; // Changed to "Image Path"
            }
        };


        // 3. Main control panel (for Row 2 of mainLayout)
        TableLayoutPanel controlPanel = new TableLayoutPanel();
        controlPanel.Dock = DockStyle.Fill; // Fill its cell in mainLayout
        controlPanel.Padding = new Padding(10);
        controlPanel.ColumnCount = 3; // Two columns for text/weight, one for buttons
        controlPanel.RowCount = 8; // Adjusted rows to 8 for the new button placement and status label
        controlPanel.BackColor = SettingsManager.GetControlBackColor(); // Explicitly set background for the panel

        // Configure column and row styles for responsive layout within controlPanel
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Left column for labels
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Middle column for input fields
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F)); // Right column for buttons

        for (int i = 0; i < controlPanel.RowCount; i++)
        {
            controlPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Fixed height for rows
        }
        controlPanel.RowStyles[1].SizeType = SizeType.AutoSize; // Allow message text row to auto-size
        controlPanel.RowStyles[6].SizeType = SizeType.Absolute; // Ensure save/load row has fixed height
        controlPanel.RowStyles[7].SizeType = SizeType.AutoSize; // Allow status label row to auto-size (moved to row 7)


        // Controls for adding new messages (New layout using TableLayoutPanel)

        // New Message Text Label
        lblText = new Label() { Text = "New Message Text:", Anchor = AnchorStyles.Left };
        controlPanel.Controls.Add(lblText, 0, 0); // Column 0, Row 0

        // New Message Textbox
        txtMessageText = new TextBox() { Multiline = true, ScrollBars = ScrollBars.Vertical, Height = 60, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom, BorderStyle = BorderStyle.FixedSingle };
        txtMessageText.BackColor = SettingsManager.GetButtonActiveColor();
        txtMessageText.ForeColor = SettingsManager.GetForeColor();
        toolTip1.SetToolTip(txtMessageText, "Enter the message text to send.");
        controlPanel.Controls.Add(txtMessageText, 1, 0); // Column 1, Row 0
        controlPanel.SetRowSpan(txtMessageText, 2); // Span two rows for larger text box

        // New Message Weight Label
        lblWeight = new Label() { Text = "New Message Weight:", Anchor = AnchorStyles.Left };
        controlPanel.Controls.Add(lblWeight, 0, 2); // Column 0, Row 2

        // New Message Weight NumericUpDown
        nudMessageWeight = new NumericUpDown() { Minimum = 1, Maximum = int.MaxValue, Value = 1000, Anchor = AnchorStyles.Left, BorderStyle = BorderStyle.FixedSingle };
        nudMessageWeight.BackColor = SettingsManager.GetButtonActiveColor();
        nudMessageWeight.ForeColor = SettingsManager.GetForeColor();
        toolTip1.SetToolTip(nudMessageWeight, "Set the weight for message selection (higher means more frequent).");
        controlPanel.Controls.Add(nudMessageWeight, 1, 2); // Column 1, Row 2

        // Image URL Label
        lblImageUrl = new Label() { Text = "Image Path:", Anchor = AnchorStyles.Left }; // New: Image URL Label
        controlPanel.Controls.Add(lblImageUrl, 0, 3); // Column 0, Row 3

        // Image URL Textbox
        txtImageUrl = new TextBox() { Anchor = AnchorStyles.Left | AnchorStyles.Right, BorderStyle = BorderStyle.FixedSingle };
        txtImageUrl.BackColor = SettingsManager.GetButtonActiveColor();
        txtImageUrl.ForeColor = SettingsManager.GetForeColor();
        toolTip1.SetToolTip(txtImageUrl, "Enter a URL or local path to an image to send.");
        controlPanel.Controls.Add(txtImageUrl, 1, 3); // Column 1, Row 3

        // Loop Count Label
        lblLoopCount = new Label() { Text = "Send Count:", Anchor = AnchorStyles.Left };
        controlPanel.Controls.Add(lblLoopCount, 0, 4); // Column 0, Row 4

        // Loop Count NumericUpDown
        nudLoopCount = new NumericUpDown() { Minimum = 1, Maximum = 99999, Value = 1, Anchor = AnchorStyles.Left };
        nudLoopCount.BackColor = SettingsManager.GetButtonActiveColor();
        nudLoopCount.ForeColor = SettingsManager.GetForeColor();
        toolTip1.SetToolTip(nudLoopCount, "Number of times to loop through sending messages.");
        controlPanel.Controls.Add(nudLoopCount, 1, 4); // Column 1, Row 4


        // Buttons

        // Add New Message Button
        btnAdd = new Button() { Text = "Add New Message", Dock = DockStyle.Fill };
        btnAdd.Click += BtnAdd_Click;
        toolTip1.SetToolTip(btnAdd, "Add a new message to the list. (Hotkey: Ctrl+N)");
        controlPanel.Controls.Add(btnAdd, 2, 0); // Column 2, Row 0

        // Edit Selected Message Button
        btnUpdate = new Button() { Text = "Edit Selected Message", Dock = DockStyle.Fill, Enabled = false }; // Initially disabled
        btnUpdate.Click += BtnUpdate_Click;
        toolTip1.SetToolTip(btnUpdate, "Edit the selected message. (Hotkey: Enter)");
        controlPanel.Controls.Add(btnUpdate, 2, 1); // Column 2, Row 1

        // Delete Selected Message Button
        btnDelete = new Button() { Text = "Delete Selected Message", Dock = DockStyle.Fill, Enabled = false }; // Initially disabled
        btnDelete.Click += BtnDelete_Click;
        toolTip1.SetToolTip(btnDelete, "Delete the selected message. (Hotkey: Del)");
        controlPanel.Controls.Add(btnDelete, 2, 2); // Column 2, Row 2

        // Duplicate Selected Message Button
        btnDuplicate = new Button() { Text = "Duplicate Selected", Dock = DockStyle.Fill, Enabled = false };
        btnDuplicate.Click += BtnDuplicate_Click;
        toolTip1.SetToolTip(btnDuplicate, "Create a copy of the selected message. (Hotkey: Ctrl+D)");
        controlPanel.Controls.Add(btnDuplicate, 2, 3); // New duplicate button position

        // Start/Stop Sending Button
        btnStartStop = new Button() { Text = "Start Sending", Dock = DockStyle.Fill };
        btnStartStop.Click += BtnStartStop_Click;
        toolTip1.SetToolTip(btnStartStop, "Start or stop sending messages to Discord. (Hotkey: F5)");
        controlPanel.Controls.Add(btnStartStop, 2, 4); // Column 2, Row 4
        controlPanel.SetRowSpan(btnStartStop, 2); // Make Start/Stop button span two rows

        // Export Messages Button
        btnSaveMessages = new Button() { Text = "Export Messages", Dock = DockStyle.Fill }; // Changed text to Export
        btnSaveMessages.Click += BtnSaveMessages_Click;
        toolTip1.SetToolTip(btnSaveMessages, "Export all messages to a JSON file. (Hotkey: Ctrl+S)");
        controlPanel.Controls.Add(btnSaveMessages, 0, 6); // Column 0, Row 6 (Moved from row 5)

        // Import Messages Button
        btnLoadMessages = new Button() { Text = "Import Messages", Dock = DockStyle.Fill }; // Changed text to Import
        btnLoadMessages.Click += BtnLoadMessages_Click;
        toolTip1.SetToolTip(btnLoadMessages, "Import messages from a JSON file. (Hotkey: Ctrl+L)");
        controlPanel.Controls.Add(btnLoadMessages, 1, 6); // Column 1, Row 6 (Moved from row 5)

        // Settings Button
        btnSettings = new Button() { Text = "Settings", Dock = DockStyle.Fill };
        btnSettings.Click += BtnSettings_Click;
        toolTip1.SetToolTip(btnSettings, "Open application settings (colors, hotkeys).");
        controlPanel.Controls.Add(btnSettings, 2, 6); // Column 2, Row 6

        // Status Label
        lblStatus = new Label() { Text = "Status: Idle", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        controlPanel.Controls.Add(lblStatus, 0, 7); // Column 0, Row 7 (Moved from row 6)
        controlPanel.SetColumnSpan(lblStatus, 3); // Span across all columns
        mainLayout.Controls.Add(controlPanel, 0, 2); // Add control panel to Row 2


        // 4. Debug log console (for Row 3 of mainLayout)
        rtbLog = new RichTextBox();
        rtbLog.Dock = DockStyle.Fill; // Set to Fill to occupy the entire row
        rtbLog.ReadOnly = true;
        rtbLog.BackColor = SettingsManager.GetButtonActiveColor();
        rtbLog.ForeColor = SettingsManager.GetForeColor();
        rtbLog.WordWrap = false; // Prevent word wrapping in log
        rtbLog.ScrollBars = RichTextBoxScrollBars.Both; // Both scrollbars for wide messages
        rtbLog.BorderStyle = BorderStyle.FixedSingle; // Added BorderStyle
        mainLayout.Controls.Add(rtbLog, 0, 3); // Add log console to Row 3

        this.Controls.Add(mainLayout); // Add the main layout panel to the form
    }

    private void ApplyTheme()
    {
        // Apply form level colors
        this.BackColor = SettingsManager.GetBackColor();
        this.ForeColor = SettingsManager.GetForeColor();

        // Apply theme to controls recursively
        ApplyThemeToControls(this.Controls);

        // Specific DataGridView styling that might not be covered by general recursion
        dgvMessages.BackgroundColor = SettingsManager.GetBackColor();
        dgvMessages.GridColor = SettingsManager.GetControlBackColor();
        dgvMessages.DefaultCellStyle.BackColor = SettingsManager.GetControlBackColor();
        dgvMessages.DefaultCellStyle.ForeColor = SettingsManager.GetForeColor();
        dgvMessages.DefaultCellStyle.SelectionBackColor = Color.FromArgb(70, 73, 79);
        dgvMessages.DefaultCellStyle.SelectionForeColor = SettingsManager.GetForeColor();
        dgvMessages.ColumnHeadersDefaultCellStyle.BackColor = SettingsManager.GetDataGridViewHeaderBackColor();
        dgvMessages.ColumnHeadersDefaultCellStyle.ForeColor = SettingsManager.GetForeColor();
        dgvMessages.RowHeadersDefaultCellStyle.BackColor = SettingsManager.GetDataGridViewHeaderBackColor();
        dgvMessages.RowHeadersDefaultCellStyle.ForeColor = SettingsManager.GetForeColor();

        // Update the log textbox and search textbox explicitly
        rtbLog.BackColor = SettingsManager.GetButtonActiveColor();
        rtbLog.ForeColor = SettingsManager.GetForeColor();
        txtSearch.BackColor = SettingsManager.GetButtonActiveColor();
        txtSearch.ForeColor = SettingsManager.GetForeColor();
    }

    private void ApplyThemeToControls(Control.ControlCollection controls)
    {
        foreach (Control control in controls)
        {
            control.ForeColor = SettingsManager.GetForeColor();

            if (control is Button button)
            {
                // Unsubscribe to avoid double-subscription if ApplyTheme is called multiple times
                button.MouseEnter -= Button_MouseEnter;
                button.MouseLeave -= Button_MouseLeave;
                button.MouseDown -= Button_MouseDown;
                button.MouseUp -= Button_MouseUp;

                // Subscribe to events for hover/active effects
                button.MouseEnter += Button_MouseEnter;
                button.MouseLeave += Button_MouseLeave;
                button.MouseDown += Button_MouseDown;
                button.MouseUp += Button_MouseUp;

                button.FlatStyle = FlatStyle.Flat; // Use Flat style for custom border
                button.BackColor = SettingsManager.GetControlBackColor();
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 90); // Subtle border
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = SettingsManager.GetButtonActiveColor();
                textBox.BorderStyle = BorderStyle.FixedSingle; // Ensure consistent border
            }
            else if (control is NumericUpDown numericUpDown)
            {
                numericUpDown.BackColor = SettingsManager.GetButtonActiveColor();
                numericUpDown.BorderStyle = BorderStyle.FixedSingle; // Ensure consistent border
            }
            else if (control is Label label)
            {
                label.BackColor = Color.Transparent;
            }
            else if (control is RichTextBox richTextBox)
            {
                richTextBox.BackColor = SettingsManager.GetButtonActiveColor();
                richTextBox.BorderStyle = BorderStyle.FixedSingle; // Ensure consistent border
            }
            else if (control is Panel || control is TableLayoutPanel)
            {
                // Only set background for panels/layout panels if they are direct containers
                // and not for the form itself which has DarkBackColor
                if (control != this) // Ensure we don't overwrite the main form's back color
                {
                    control.BackColor = SettingsManager.GetControlBackColor();
                }
            }
            else if (control is TrackBar trackBar) // New: Apply theme to TrackBar
            {
                trackBar.BackColor = SettingsManager.GetControlBackColor();
                trackBar.ForeColor = SettingsManager.GetForeColor(); // For tick marks etc.
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.BackColor = Color.Transparent; // Checkboxes usually have transparent background
            }
            // Recursively apply theme to child controls
            if (control.HasChildren)
            {
                ApplyThemeToControls(control.Controls);
            }
        }
    }

    // Button event handlers for hover/active effects
    private void Button_MouseEnter(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture) // Only change if not currently pressed (captured)
        {
            button.BackColor = SettingsManager.GetButtonHoverColor();
        }
    }

    private void Button_MouseLeave(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture) // Only revert if not currently pressed
        {
            button.BackColor = SettingsManager.GetControlBackColor();
        }
    }

    private void Button_MouseDown(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            button.BackColor = SettingsManager.GetButtonActiveColor();
        }
    }

    private void Button_MouseUp(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            if (button.ClientRectangle.Contains(button.PointToClient(Cursor.Position)))
            {
                // Mouse is still over the button, revert to hover color
                button.BackColor = SettingsManager.GetButtonHoverColor();
            }
            else
            {
                // Mouse left the button, revert to normal color
                button.BackColor = SettingsManager.GetControlBackColor();
            }
        }
    }


    private void InitializeData()
    {
        // Define default messages. Only the ":3" message is hardcoded now.
        // The ImageUrl is set to an empty string by default.
        allMessages = new List<Message>
        {
            new Message(":3", 6000, "")
        };

        // Initialize available messages with a copy of all messages
        availableMessages = new List<Message>(allMessages);

        // Set up BindingSource to manage the DataGridView's data
        messageBindingSource = new BindingSource();
        messageBindingSource.DataSource = allMessages;
        dgvMessages.DataSource = messageBindingSource;

        inputSimulator = new InputSimulator(); // Used for simulating keyboard input
        random = new Random(); // Random number generator for message selection and delays
    }

    // Event handler for adding a new message
    private void BtnAdd_Click(object sender, EventArgs e)
    {
        string text = txtMessageText.Text.Trim();
        int weight = (int)nudMessageWeight.Value;
        string imageUrl = txtImageUrl.Text.Trim(); // Get image URL

        if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(imageUrl)) // Message needs at least text or image
        {
            Log("Message text and Image URL cannot both be empty.");
            MessageBox.Show("Message text and Image URL cannot both be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Message newMessage = new Message(text, weight, imageUrl); // Pass imageUrl to constructor
        allMessages.Add(newMessage);
        messageBindingSource.ResetBindings(false); // Refresh DataGridView
        availableMessages = new List<Message>(allMessages); // Reset available messages when adding new
        txtMessageText.Clear();
        nudMessageWeight.Value = 1000;
        txtImageUrl.Clear(); // Clear image URL field
        Log($"Added new message: Text='{text}', Weight={weight}, ImageUrl='{imageUrl}'");
        // No automatic save here, rely on explicit Save/Load
    }

    // Event handler for updating a selected message
    private void BtnUpdate_Click(object sender, EventArgs e)
    {
        if (dgvMessages.SelectedRows.Count == 0)
        {
            Log("No message selected for update.");
            MessageBox.Show("Please select a message to update.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        EditSelectedMessage();
    }

    // New method for editing selected message, callable by button and hotkey
    private void EditSelectedMessage()
    {
        if (dgvMessages.SelectedRows.Count == 0) return; // Should be handled by calling context

        // Get the selected message object from the DataGridView's data source
        Message selectedMessage = (Message)dgvMessages.SelectedRows[0].DataBoundItem;
        Log($"Attempting to edit message: Text='{selectedMessage.Text}'");

        // Create and show the new EditMessageForm
        using (EditMessageForm editForm = new EditMessageForm(selectedMessage))
        {
            // Pass current theme colors to the edit form from settings manager
            editForm.SetThemeColors(
                SettingsManager.GetBackColor(),
                SettingsManager.GetForeColor(),
                SettingsManager.GetControlBackColor(),
                SettingsManager.GetButtonActiveColor(),
                SettingsManager.GetDataGridViewHeaderBackColor()
            );

            // If the user clicks 'Save' in the edit form
            if (editForm.ShowDialog() == DialogResult.OK)
            {
                // The 'selectedMessage' object itself has been updated by EditMessageForm
                // So we just need to refresh the DataGridView and reset available messages.
                messageBindingSource.ResetBindings(false); // Refresh DataGridView to reflect changes
                availableMessages = new List<Message>(allMessages); // Reset available messages when updating
                Log($"Message updated successfully: Text='{selectedMessage.Text}', Weight={selectedMessage.Weight}, ImageUrl='{selectedMessage.ImageUrl}'");
                MessageBox.Show("Message updated successfully.", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // No automatic save here, rely on explicit Save/Load
            }
            else
            {
                Log("Message update cancelled.");
            }
        }
    }


    // Event handler for deleting a selected message
    private void BtnDelete_Click(object sender, EventArgs e)
    {
        if (dgvMessages.SelectedRows.Count == 0)
        {
            Log("No message selected for deletion.");
            MessageBox.Show("Please select a message to delete.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        DeleteSelectedMessage();
    }

    // New method for deleting selected message, callable by button and hotkey
    private void DeleteSelectedMessage()
    {
        if (dgvMessages.SelectedRows.Count == 0) return; // Should be handled by calling context

        Message selectedMessage = (Message)dgvMessages.SelectedRows[0].DataBoundItem;
        Log($"Attempting to delete message: Text='{selectedMessage.Text}'");

        // Confirmation dialog for deletion
        DialogResult confirmResult = MessageBox.Show(
            "Are you sure you want to delete the selected message?",
            "Confirm Delete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirmResult == DialogResult.Yes)
        {
            allMessages.Remove(selectedMessage);
            messageBindingSource.ResetBindings(false); // Refresh DataGridView
            availableMessages = new List<Message>(allMessages); // Reset available messages when deleting
            // After deletion, clear text boxes for new message entry and disable update/delete buttons
            txtMessageText.Clear();
            nudMessageWeight.Value = 1000;
            txtImageUrl.Clear(); // Clear image URL field
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            btnDuplicate.Enabled = false; // Disable duplicate button too
            Log($"Message deleted successfully: Text='{selectedMessage.Text}'");
            MessageBox.Show("Message deleted successfully.", "Delete Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            Log("Message deletion cancelled.");
        }
    }

    // New: Event handler for duplicating a selected message
    private void BtnDuplicate_Click(object sender, EventArgs e)
    {
        if (dgvMessages.SelectedRows.Count == 0)
        {
            Log("No message selected for duplication.");
            MessageBox.Show("Please select a message to duplicate.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Message selectedMessage = (Message)dgvMessages.SelectedRows[0].DataBoundItem;
        // Create a new message with the same content
        Message duplicatedMessage = new Message(selectedMessage.Text, selectedMessage.Weight, selectedMessage.ImageUrl);
        allMessages.Add(duplicatedMessage);
        messageBindingSource.ResetBindings(false); // Refresh DataGridView
        availableMessages = new List<Message>(allMessages); // Reset available messages
        Log($"Duplicated message: Text='{selectedMessage.Text}'");
        MessageBox.Show("Message duplicated successfully.", "Duplicate Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }


    // Event handler for DataGridView selection changes
    private void DgvMessages_SelectionChanged(object sender, EventArgs e)
    {
        // Enable/disable update/delete/duplicate buttons based on whether a row is selected
        bool rowSelected = dgvMessages.SelectedRows.Count > 0;
        btnUpdate.Enabled = rowSelected;
        btnDelete.Enabled = rowSelected;
        btnDuplicate.Enabled = rowSelected; // Enable/disable duplicate button

        // Clear the new message input fields when a message is selected in the grid
        // to clearly separate "add" functionality from "edit" functionality.
        txtMessageText.Clear();
        nudMessageWeight.Value = 1000;
        txtImageUrl.Clear(); // Clear image URL field
        // Log("DataGridView selection changed. Buttons updated."); // Too chatty for general use
    }

    // New: Handle double-click on DataGridView cell to edit message
    private void DgvMessages_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0 && e.RowIndex < dgvMessages.Rows.Count)
        {
            dgvMessages.Rows[e.RowIndex].Selected = true; // Ensure row is selected
            EditSelectedMessage();
        }
    }

    // New: Handle search text box changes to filter messages
    private void TxtSearch_TextChanged(object sender, EventArgs e)
    {
        string searchText = txtSearch.Text.Trim().ToLower();
        if (string.IsNullOrEmpty(searchText))
        {
            messageBindingSource.DataSource = allMessages; // Show all messages if search is empty
            Log("Search cleared. Displaying all messages.");
        }
        else
        {
            var filteredList = allMessages.Where(m =>
                m.Text.ToLower().Contains(searchText) ||
                m.ImageUrl.ToLower().Contains(searchText)
            ).ToList();
            messageBindingSource.DataSource = filteredList; // Update DGV with filtered list
            Log($"Filtered messages by: '{searchText}'. Found {filteredList.Count} results.");
        }
        messageBindingSource.ResetBindings(false); // Crucial to refresh after filtering
    }

    // New: Hotkey handling for the main form
    private void MainForm_KeyDown(object sender, KeyEventArgs e)
    {
        // Helper to parse hotkey string to Keys enum, handling potential parsing errors
        Keys ParseHotkey(string hotkeyString)
        {
            if (Enum.TryParse(hotkeyString, out Keys hotkey))
            {
                return hotkey;
            }
            return Keys.None; // Return None if parsing fails
        }

        bool ctrlPressed = e.Control;

        // Hotkey for Export (Ctrl+S)
        if (ctrlPressed && e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.SaveHotkey))
        {
            BtnSaveMessages_Click(this, EventArgs.Empty);
            e.Handled = true;
        }
        // Hotkey for Import (Ctrl+L)
        else if (ctrlPressed && e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.LoadHotkey))
        {
            BtnLoadMessages_Click(this, EventArgs.Empty);
            e.Handled = true;
        }
        // Hotkey for Add New Message (Ctrl+N)
        else if (ctrlPressed && e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.AddHotkey))
        {
            BtnAdd_Click(this, EventArgs.Empty);
            e.Handled = true;
        }
        // Hotkey for Duplicate Selected Message (Ctrl+D)
        else if (ctrlPressed && e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.DuplicateHotkey))
        {
            BtnDuplicate_Click(this, EventArgs.Empty);
            e.Handled = true;
        }
        // Hotkey for Delete Selected Message (Del)
        else if (e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.DeleteHotkey))
        {
            DeleteSelectedMessage();
            e.Handled = true;
        }
        // Hotkey for Edit Selected Message (Enter on DGV)
        else if (e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.EditHotkey) && dgvMessages.Focused && dgvMessages.SelectedRows.Count > 0)
        {
            EditSelectedMessage();
            e.Handled = true;
        }
        // Hotkey for Start/Stop Sending (F5)
        else if (e.KeyCode == ParseHotkey(SettingsManager.CurrentSettings.SendHotkey))
        {
            BtnStartStop_Click(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    // DataGridView Drag & Drop implementation
    private void DgvMessages_MouseDown(object sender, MouseEventArgs e)
    {
        // Get the index of the item the mouse is currently over.
        rowIndexFromMouseDown = dgvMessages.HitTest(e.X, e.Y).RowIndex;
        if (rowIndexFromMouseDown != -1)
        {
            // If it's a valid row, remember its position for drag operation.
            Size dragSize = SystemInformation.DragSize;
            dragBoxFromMouseDown = new Rectangle(new Point(e.X - (dragSize.Width / 2),
                                                           e.Y - (dragSize.Height / 2)),
                                                 dragSize);
        }
        else
        {
            // Reset drag box if no valid row is clicked.
            dragBoxFromMouseDown = Rectangle.Empty;
        }
    }

    private void DgvMessages_MouseMove(object sender, MouseEventArgs e)
    {
        if ((e.Button & MouseButtons.Left) == MouseButtons.Left && dragBoxFromMouseDown != Rectangle.Empty)
        {
            if (!dragBoxFromMouseDown.Contains(e.X, e.Y))
            {
                // Start the drag-and-drop operation
                dgvMessages.DoDragDrop(dgvMessages.Rows[rowIndexFromMouseDown], DragDropEffects.Move);
            }
        }
    }

    private void DgvMessages_DragOver(object sender, DragEventArgs e)
    {
        e.Effect = DragDropEffects.Move;
    }

    private void DgvMessages_DragDrop(object sender, DragEventArgs e)
    {
        if (e.Effect == DragDropEffects.Move)
        {
            Point clientPoint = dgvMessages.PointToClient(new Point(e.X, e.Y));
            rowIndexOfItemUnderMouseToDrop = dgvMessages.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            if (rowIndexOfItemUnderMouseToDrop == -1 || rowIndexOfItemUnderMouseToDrop == rowIndexFromMouseDown)
            {
                // Dropped outside grid or on the same row, do nothing
                return;
            }

            // Get the dragged row
            DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;

            if (rowToMove != null)
            {
                // Get the actual Message object
                Message messageToMove = (Message)rowToMove.DataBoundItem;

                // Remove from current position in both allMessages and availableMessages
                allMessages.RemoveAt(rowIndexFromMouseDown);
                availableMessages.Remove(messageToMove); // Ensure it's removed from available if present

                // Insert into new position in allMessages
                allMessages.Insert(rowIndexOfItemUnderMouseToDrop, messageToMove);

                // Re-populate availableMessages (simplest way to ensure order consistency)
                availableMessages = new List<Message>(allMessages);

                // Refresh the DataGridView
                messageBindingSource.ResetBindings(false);

                Log($"Message moved from row {rowIndexFromMouseDown} to {rowIndexOfItemUnderMouseToDrop}.");

                // Select the moved row after reordering and refreshing
                if (rowIndexOfItemUnderMouseToDrop >= 0 && rowIndexOfItemUnderMouseToDrop < dgvMessages.Rows.Count)
                {
                    dgvMessages.ClearSelection();
                    dgvMessages.Rows[rowIndexOfItemUnderMouseToDrop].Selected = true;
                }
            }
        }
    }


    // Event handler for Start/Stop button
    private async void BtnStartStop_Click(object sender, EventArgs e)
    {
        if (!isSending)
        {
            int maxLoops = (int)nudLoopCount.Value;
            if (maxLoops <= 0)
            {
                Log("Error: Loop count must be at least 1.");
                MessageBox.Show("Loop count must be at least 1.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (allMessages.Count == 0)
            {
                Log("Error: No messages available to send. Please add messages first.");
                MessageBox.Show("No messages available to send. Please add messages first.", "No Messages", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Log("Starting message sending process...");
            btnStartStop.Text = "Stop Sending";
            lblStatus.Text = "Status: Sending messages...";
            isSending = true;
            cancellationTokenSource = new CancellationTokenSource();

            // Disable edit controls while sending is active
            btnAdd.Enabled = false;
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            btnDuplicate.Enabled = false;
            dgvMessages.Enabled = false;
            btnSaveMessages.Enabled = false;
            btnLoadMessages.Enabled = false;
            btnSettings.Enabled = false; // Disable settings button during sending
            nudLoopCount.Enabled = false;

            // Run the sending logic in a separate task to keep the UI responsive
            sendingTask = Task.Run(async () => await SendMessagesLoop(cancellationTokenSource.Token, maxLoops));
        }
        else
        {
            Log("Stopping message sending process...");
            btnStartStop.Text = "Start Sending";
            lblStatus.Text = "Status: Stopping...";
            isSending = false;
            cancellationTokenSource?.Cancel(); // Request cancellation

            // Wait for the task to complete its current iteration or finish
            if (sendingTask != null)
            {
                await sendingTask; // Wait for the task to gracefully finish
            }
            Log("Message sending process stopped.");
            // UI re-enabling is now handled in the finally block of SendMessagesLoop
        }
    }

    /// <summary>
    /// Sends an image to the foreground Discord window by copying it to the clipboard and simulating paste.
    /// </summary>
    /// <param name="imageUrl">The URL or local path of the image to send.</param>
    /// <param name="cancellationToken">Cancellation token to stop the operation.</param>
    /// <returns>True if image was sent, false otherwise.</returns>
    private async Task<bool> SendImageFromUrl(string imageUrl, CancellationToken cancellationToken)
    {
        Image imageToCopy = null;
        Log($"Attempting to send image from URL/path: {imageUrl}");
        try
        {
            if (Uri.TryCreate(imageUrl, UriKind.Absolute, out Uri uriResult) &&
                (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps))
            {
                Log($"Downloading image from URL: {imageUrl}");
                // It's a web URL, download the image
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(15); // Increased timeout for image download
                    byte[] imageBytes = await client.GetByteArrayAsync(uriResult, cancellationToken);
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        imageToCopy = Image.FromStream(ms);
                    }
                    Log($"Image downloaded successfully. Size: {imageBytes.Length} bytes.");
                }
            }
            else if (File.Exists(imageUrl))
            {
                Log($"Loading image from local file: {imageUrl}");
                // It's a local file path, load the image
                imageToCopy = Image.FromFile(imageUrl);
                Log("Image loaded from file successfully.");
            }
            else
            {
                Log($"Error: Invalid image URL or file path: '{imageUrl}'");
                this.Invoke((MethodInvoker)delegate
                {
                    lblStatus.Text = $"Status: Invalid image URL/path: {imageUrl.Substring(0, Math.Min(50, imageUrl.Length))}...";
                });
                return false;
            }

            if (imageToCopy != null)
            {
                // Clipboard operations must be on the UI thread
                // Implement retry logic for clipboard operations
                const int maxRetries = 5;
                for (int i = 0; i < maxRetries; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            Clipboard.SetImage(imageToCopy);
                        });
                        Log($"Image successfully copied to clipboard (Attempt {i + 1}/{maxRetries}).");
                        break; // Exit retry loop on success
                    }
                    catch (ExternalException ex) // This is typically a COMException when clipboard is locked
                    {
                        Log($"Warning: Failed to copy image to clipboard (Attempt {i + 1}/{maxRetries}): {ex.Message}. Retrying in 100ms...");
                        if (i < maxRetries - 1)
                        {
                            await Task.Delay(100, cancellationToken); // Wait a bit before retrying
                            this.Invoke((MethodInvoker)delegate
                            {
                                Clipboard.Clear(); // Attempt to clear a potentially locked clipboard
                            });
                            await Task.Delay(50, cancellationToken); // Give time for clear
                        }
                        else
                        {
                            Log($"Error: Failed to copy image to clipboard after {maxRetries} attempts: {ex.Message}");
                            this.Invoke((MethodInvoker)delegate
                            {
                                lblStatus.Text = $"Status: Failed to copy image to clipboard: {ex.Message.Substring(0, Math.Min(50, ex.Message.Length))}";
                            });
                            return false; // All retries failed
                        }
                    }
                }

                await Task.Delay(ClipboardClearDelayMs, cancellationToken); // Give clipboard a moment after setting image

                Log("Simulating Ctrl+V (paste) for image.");
                // Simulate Ctrl+V (paste)
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                await Task.Delay(ImagePasteDelayMs, cancellationToken); // Give time for paste to register

                this.Invoke((MethodInvoker)delegate
                {
                    Clipboard.Clear(); // Clear the clipboard after pasting to avoid side effects
                });
                Log("Image paste simulated and clipboard cleared.");
                return true;
            }
        }
        catch (OperationCanceledException)
        {
            Log("Image sending cancelled by user.");
            throw; // Re-throw to be caught by the main loop's cancellation handler
        }
        catch (HttpRequestException ex)
        {
            Log($"HTTP Error downloading image from URL: {ex.Message}");
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = $"Status: Failed to download image from URL: {ex.Message.Substring(0, Math.Min(50, imageUrl.Length))}";
            });
        }
        catch (FileNotFoundException)
        {
            Log($"File Not Found error for image: {imageUrl}");
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = $"Status: Image file not found: {imageUrl.Substring(0, Math.Min(50, imageUrl.Length))}";
            });
        }
        catch (ArgumentException ex) // For invalid image formats from FromStream/FromFile
        {
            Log($"Argument Error (Invalid image format): {ex.Message}");
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = $"Status: Invalid image format: {ex.Message.Substring(0, Math.Min(50, ex.Message.Length))}";
            });
        }
        catch (Exception ex)
        {
            Log($"Unhandled Error during image sending: {ex.GetType().Name}: {ex.Message}");
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = $"Status: Error sending image: {ex.Message.Substring(0, Math.Min(50, ex.Message.Length))}";
            });
        }
        finally
        {
            imageToCopy?.Dispose(); // Dispose the image object if it was loaded
            Log("Image object disposed.");
        }
        return false;
    }


    // The core message sending loop, now asynchronous and takes maxLoops as argument
    private async Task SendMessagesLoop(CancellationToken cancellationToken, int maxLoops)
    {
        int currentLoop = 0; // Track current loop count
        Log($"SendMessagesLoop started for {maxLoops} loops.");
        try // Outer try-catch for OperationCanceledException
        {
            while (!cancellationToken.IsCancellationRequested && currentLoop < maxLoops)
            {
                Log($"Current loop: {currentLoop + 1}/{maxLoops}");

                // Conditionally reset message pool based on AllowDuplicateMessages setting
                if (!SettingsManager.CurrentSettings.AllowDuplicateMessages && availableMessages.Count == 0)
                {
                    Log("All available unique messages sent in current cycle. Resetting message pool.");
                    this.Invoke((MethodInvoker)delegate {
                        lblStatus.Text = $"Status: All unique messages sent. Resetting pool. (Loop {currentLoop + 1}/{maxLoops})";
                    });
                    availableMessages = new List<Message>(allMessages); // Refill with all original messages
                }
                else if (SettingsManager.CurrentSettings.AllowDuplicateMessages)
                {
                    // If duplicates are allowed, the availableMessages list is always just a copy of allMessages
                    // and we don't remove from it. So we just ensure it's up to date.
                    if (availableMessages.Count != allMessages.Count)
                    {
                        availableMessages = new List<Message>(allMessages);
                    }
                }

                if (availableMessages.Count == 0) // No messages in the master list or available list
                {
                    Log("Error: No messages available in master list after reset. Exiting loop.");
                    this.Invoke((MethodInvoker)delegate
                    {
                        lblStatus.Text = "Status: No messages available to send.";
                        MessageBox.Show("No messages available to send. Please add messages first.", "No Messages", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    });
                    break; // Exit loop if no messages at all
                }

                await Task.Delay(500, cancellationToken); // This can throw TaskCanceledException

                // Find all running Discord processes
                Process[] discordProcesses = Process.GetProcessesByName("Discord");
                IntPtr discordWindow = IntPtr.Zero;

                foreach (var proc in discordProcesses)
                {
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                        discordWindow = proc.MainWindowHandle;
                        break;
                    }
                }

                if (discordWindow != IntPtr.Zero)
                {
                    Log("Discord window found. Forcing focus.");
                    ForceFocusWindow(discordWindow);
                    await Task.Delay(FocusDelayMs, cancellationToken); // Use Task.Delay for non-blocking wait

                    string messageTextToSend = "";
                    string imageUrlToSend = "";
                    int totalWeightAvailable = availableMessages.Sum(m => m.Weight);

                    if (totalWeightAvailable == 0)
                    {
                        Log("Warning: No messages with positive weight available in current pool. Skipping loop.");
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: No messages with weight available to send! (Loop {currentLoop + 1}/{maxLoops})";
                        });
                        await Task.Delay(5000, cancellationToken); // Wait and re-check
                        currentLoop++; // Increment loop even if no message was sent
                        continue;
                    }

                    int randomNumber = random.Next(totalWeightAvailable);
                    Message selectedMessage = null;
                    foreach (var msg in availableMessages)
                    {
                        if (randomNumber < msg.Weight)
                        {
                            selectedMessage = msg;
                            break;
                        }
                        randomNumber -= msg.Weight;
                    }

                    if (selectedMessage != null)
                    {
                        messageTextToSend = selectedMessage.Text;
                        imageUrlToSend = selectedMessage.ImageUrl;
                        Log($"Selected message: Text='{messageTextToSend}', ImageUrl='{imageUrlToSend}' (Weight: {selectedMessage.Weight})");

                        // Highlight the row currently being sent (on UI thread)
                        int selectedMessageIndex = allMessages.IndexOf(selectedMessage);
                        if (selectedMessageIndex >= 0 && selectedMessageIndex < dgvMessages.Rows.Count)
                        {
                            this.Invoke((MethodInvoker)delegate
                            {
                                if (_activeMessageRow != null) // Reset previous highlight
                                {
                                    _activeMessageRow.DefaultCellStyle.BackColor = SettingsManager.GetControlBackColor(); // Default row color
                                }
                                _activeMessageRow = dgvMessages.Rows[selectedMessageIndex];
                                _activeMessageRow.DefaultCellStyle.BackColor = SettingsManager.GetHighlightColor(); // Highlight color
                            });
                        }

                        // Only remove from available messages if AllowDuplicateMessages is false
                        if (!SettingsManager.CurrentSettings.AllowDuplicateMessages)
                        {
                            this.Invoke((MethodInvoker)delegate {
                                availableMessages.Remove(selectedMessage);
                            });
                        }
                    }
                    else
                    {
                        Log("Error: Failed to select a message. Retrying...");
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Error selecting message. Retrying... (Loop {currentLoop + 1}/{maxLoops})";
                        });
                        await Task.Delay(InputDelayMs * 2, cancellationToken);
                        currentLoop++; // Increment loop even on error
                        continue;
                    }

                    bool imageSentSuccessfully = false;
                    bool textSentSuccessfully = false;

                    // Try to send image first if available
                    if (!string.IsNullOrEmpty(imageUrlToSend))
                    {
                        Log($"Sending image: {imageUrlToSend}");
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Preparing to send image... (Loop {currentLoop + 1}/{maxLoops})";
                        });
                        imageSentSuccessfully = await SendImageFromUrl(imageUrlToSend, cancellationToken);
                        if (imageSentSuccessfully)
                        {
                            this.Invoke((MethodInvoker)delegate {
                                lblStatus.Text = $"Status: Image sent. (Loop {currentLoop + 1}/{maxLoops})";
                            });
                        }
                        else
                        {
                            Log("Image sending failed. See earlier logs for details. Continuing with text if any.");
                            this.Invoke((MethodInvoker)delegate {
                                lblStatus.Text += " Continuing with text message if any.";
                            });
                            await Task.Delay(InputDelayMs * 2, cancellationToken);
                        }
                    }

                    // Send message text if it exists, ensuring new line if image was also sent
                    if (!string.IsNullOrEmpty(messageTextToSend))
                    {
                        Log($"Sending text message: '{messageTextToSend}'");
                        // Add a newline if an image was just sent AND it was successfully sent
                        if (imageSentSuccessfully)
                        {
                            Log("Simulating Shift+Enter for new line after image.");
                            inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.RETURN);
                            await Task.Delay(InputDelayMs, cancellationToken);
                        }

                        string[] lines = messageTextToSend.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < lines.Length; i++)
                        {
                            inputSimulator.Keyboard.TextEntry(lines[i]);
                            await Task.Delay(InputDelayMs, cancellationToken);

                            if (i < lines.Length - 1)
                            {
                                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.RETURN);
                                await Task.Delay(InputDelayMs, cancellationToken);
                            }
                        }
                        textSentSuccessfully = true;
                        Log("Text message typed.");
                    }

                    // Only press RETURN if either text or image was sent successfully
                    if (textSentSuccessfully || imageSentSuccessfully)
                    {
                        Log("Simulating Enter to send message.");
                        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN); // Send the message
                        await Task.Delay(InputDelayMs, cancellationToken);
                        Log("Message sent successfully.");
                    }
                    else // If neither text nor image could be sent
                    {
                        Log("Warning: No content (text or image) was successfully sent for this message.");
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Failed to send message (no content or errors). (Loop {currentLoop + 1}/{maxLoops})";
                        });
                        await Task.Delay(InputDelayMs * 2, cancellationToken);
                    }


                    this.Invoke((MethodInvoker)delegate {
                        string sentContentSummary = "";
                        if (!string.IsNullOrEmpty(messageTextToSend)) sentContentSummary += messageTextToSend.Replace("\n", "\\n");
                        if (!string.IsNullOrEmpty(imageUrlToSend))
                        {
                            if (!string.IsNullOrEmpty(sentContentSummary)) sentContentSummary += " | Image: ";
                            else sentContentSummary += "Image: ";
                            sentContentSummary += Path.GetFileName(imageUrlToSend); // Show just filename/last part of URL
                        }
                        lblStatus.Text = $"Status: Sent: {sentContentSummary.Substring(0, Math.Min(50, sentContentSummary.Length))}... (Loop {currentLoop + 1}/{maxLoops})"; // Truncate for status bar
                    });
                }
                else
                {
                    Log("Discord window not found. Waiting...");
                    this.Invoke((MethodInvoker)delegate {
                        lblStatus.Text = "Status: Discord window not found! Please ensure Discord is running.";
                    });
                    // If Discord isn't found, wait longer before re-checking.
                    await Task.Delay(5000, cancellationToken); // This can throw TaskCanceledException
                }
                currentLoop++; // Increment loop count after each attempt (successful or not)

                int delay = random.Next(1000, 7001);
                Log($"Waiting for {delay}ms before next message.");
                await Task.Delay(delay, cancellationToken);
            }
        }
        catch (OperationCanceledException) // Catch specific cancellation exception
        {
            Log("SendMessagesLoop cancelled by user.");
            // This block will execute when cancellation is requested and an awaitable operation is cancelled.
            // The task will now exit gracefully without throwing an unhandled exception.
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = "Status: Stopped by user cancellation.";
            });
        }
        catch (Exception ex) // Catch any other unexpected exceptions that might occur during the loop
        {
            Log($"Unhandled error in SendMessagesLoop: {ex.GetType().Name}: {ex.Message}\nStackTrace: {ex.StackTrace}");
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = $"Status: An unexpected error occurred: {ex.Message}";
                MessageBox.Show($"An unexpected error occurred during message sending: {ex.Message}", "Application Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }
        finally // This block always executes when the task finishes (either by completion, cancellation, or error)
        {
            this.Invoke((MethodInvoker)delegate {
                if (!isSending) // Only update if not already actively sending
                {
                    lblStatus.Text = "Status: Idle";
                }
                else if (currentLoop >= maxLoops && !cancellationToken.IsCancellationRequested)
                {
                    lblStatus.Text = $"Status: Finished sending {maxLoops} messages.";
                }

                btnStartStop.Text = "Start Sending";
                isSending = false; // Ensure this is reset

                // Re-enable edit controls after sending stops
                btnAdd.Enabled = true;
                btnUpdate.Enabled = (dgvMessages.SelectedRows.Count > 0); // Re-enable based on selection
                btnDelete.Enabled = (dgvMessages.SelectedRows.Count > 0); // Re-enable based on selection
                btnDuplicate.Enabled = (dgvMessages.SelectedRows.Count > 0);
                dgvMessages.Enabled = true;
                btnSaveMessages.Enabled = true;
                btnLoadMessages.Enabled = true;
                btnSettings.Enabled = true; // Re-enable settings button
                nudLoopCount.Enabled = true;

                // Reset highlight if any
                if (_activeMessageRow != null)
                {
                    _activeMessageRow.DefaultCellStyle.BackColor = SettingsManager.GetControlBackColor();
                    _activeMessageRow = null;
                }
                Log("SendMessagesLoop finished. UI controls re-enabled.");
            });
        }
    }

    /// <summary>
    /// Saves the current list of messages to a user-selected JSON file.
    /// </summary>
    private void BtnSaveMessages_Click(object sender, EventArgs e)
    {
        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
        {
            saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            saveFileDialog.Title = "Export Messages to JSON File";
            saveFileDialog.FileName = "messages.json"; // Default file name

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string json = JsonConvert.SerializeObject(allMessages, Formatting.Indented);
                    File.WriteAllText(saveFileDialog.FileName, json);
                    Log($"Messages exported to {Path.GetFileName(saveFileDialog.FileName)}.");
                    lblStatus.Text = $"Status: Messages exported to {Path.GetFileName(saveFileDialog.FileName)}.";
                }
                catch (Exception ex)
                {
                    Log($"Error exporting messages to {saveFileDialog.FileName}: {ex.Message}");
                    lblStatus.Text = $"Error exporting messages: {ex.Message}";
                    MessageBox.Show($"Error exporting messages: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Log("Message export cancelled.");
                lblStatus.Text = "Status: Export cancelled."; // Changed to "Import cancelled" as it's a save operation.
            }
        }
    }

    /// <summary>
    /// Loads messages from a user-selected JSON file into the application.
    /// </summary>
    private void BtnLoadMessages_Click(object sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.Title = "Import Messages from JSON File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string json = File.ReadAllText(openFileDialog.FileName);
                    List<Message> loadedMessages = JsonConvert.DeserializeObject<List<Message>>(json);

                    // Filter out messages that don't have the new ImageUrl property
                    // This handles loading older JSON formats gracefully
                    var filteredMessages = loadedMessages.Select(m => new Message(m.Text, m.Weight, m.GetType().GetProperty("ImageUrl") != null ? m.ImageUrl : "")).ToList();

                    if (filteredMessages != null && filteredMessages.Any())
                    {
                        allMessages.Clear();
                        allMessages.AddRange(filteredMessages);
                        messageBindingSource.ResetBindings(false); // Refresh DataGridView
                        availableMessages = new List<Message>(allMessages); // Reset available messages after loading
                        Log($"Messages imported from {Path.GetFileName(openFileDialog.FileName)}.");
                        lblStatus.Text = $"Status: Messages imported from {Path.GetFileName(openFileDialog.FileName)}.";
                    }
                    else
                    {
                        Log($"Warning: Imported file '{openFileDialog.FileName}' was empty or invalid.");
                        lblStatus.Text = "Status: Imported file was empty or invalid.";
                        MessageBox.Show("The selected file contains no messages or is in an invalid format.", "Import Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error importing messages from {openFileDialog.FileName}: {ex.Message}");
                    lblStatus.Text = $"Error importing messages: {ex.Message}";
                    MessageBox.Show($"Error importing messages: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Log("Message import cancelled.");
                lblStatus.Text = "Status: Import cancelled.";
            }
        }
    }

    // New: Event handler for opening the settings form
    private void BtnSettings_Click(object sender, EventArgs e)
    {
        using (SettingsForm settingsForm = new SettingsForm(SettingsManager.CurrentSettings))
        {
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                // Settings were saved, reload and reapply theme
                SettingsManager.LoadSettings(); // Re-load to ensure we have the latest from disk
                ApplyTheme(); // Re-apply theme to all controls
                Log("Settings updated and theme reapplied.");
                MessageBox.Show("Settings updated successfully. You may need to restart the application for some changes (e.g., admin privileges) to take full effect.", "Settings Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                Log("Settings changes cancelled.");
            }
        }
    }

    // Main entry point for the application
    [STAThread] // Required for SendKeys and Windows Forms
    static void Main()
    {
        // Admin privileges check first
        if (!IsRunningAsAdmin())
        {
            RelaunchAsAdmin();
            return;
        }

        // Run the Windows Forms application
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

// New Form for editing individual messages
public class EditMessageForm : Form
{
    private TextBox txtMessageText;
    private NumericUpDown nudMessageWeight;
    private TextBox txtImageUrl; // New: for editing image URL
    private Button btnSave;
    private Button btnCancel;

    // Declared as fields to allow theming
    private Label lblText;
    private Label lblWeight;
    private Label lblImageUrl; // New: label for image URL

    public Message EditedMessage { get; private set; }

    // Instance fields for Dark Mode Colors - now instance-specific
    private Color _backColor;
    private Color _foreColor;
    private Color _controlBackColor;
    private Color _buttonActiveColor;
    private Color _dataGridViewHeaderBackColor;


    public EditMessageForm(Message messageToEdit)
    {
        EditedMessage = messageToEdit; // Hold reference to the original message object
        InitializeComponent();
        LoadMessageData();
    }

    // Method to set theme colors from MainForm and then apply them
    public void SetThemeColors(Color backColor, Color foreColor, Color controlBackColor, Color buttonActiveColor, Color dataGridViewHeaderBackColor)
    {
        _backColor = backColor;
        _foreColor = foreColor;
        _controlBackColor = controlBackColor;
        _buttonActiveColor = buttonActiveColor;
        _dataGridViewHeaderBackColor = dataGridViewHeaderBackColor;

        ApplyTheme(); // Apply theme after setting colors
    }

    private void InitializeComponent()
    {
        // Form Properties
        this.Text = "Edit Message";
        this.Size = new System.Drawing.Size(500, 300);
        this.MinimumSize = new System.Drawing.Size(350, 250); // Minimum size for the edit form
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.Sizable; // Make the form resizable

        // Message Textbox Label
        lblText = new Label() { Text = "Message Text:", Location = new System.Drawing.Point(10, 10), AutoSize = true };
        this.Controls.Add(lblText);

        // Message Textbox
        txtMessageText = new TextBox()
        {
            Location = new System.Drawing.Point(10, 30),
            Size = new System.Drawing.Size(460, 100), // Adjusted height for more controls
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            BorderStyle = BorderStyle.FixedSingle
        };
        // Explicitly set colors for text boxes
        txtMessageText.BackColor = _buttonActiveColor;
        txtMessageText.ForeColor = _foreColor;
        this.Controls.Add(txtMessageText);


        // Message Weight NumericUpDown Label
        lblWeight = new Label() { Text = "Weight:", Location = new System.Drawing.Point(10, 140), AutoSize = true }; // Shifted down
        this.Controls.Add(lblWeight);

        // Message Weight NumericUpDown
        nudMessageWeight = new NumericUpDown()
        {
            Location = new System.Drawing.Point(70, 137),
            Size = new System.Drawing.Size(100, 20),
            Minimum = 1,
            Maximum = int.MaxValue,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BorderStyle = BorderStyle.FixedSingle
        };
        // Explicitly set colors for numeric up-down
        nudMessageWeight.BackColor = _buttonActiveColor;
        nudMessageWeight.ForeColor = _foreColor;
        this.Controls.Add(nudMessageWeight);

        // Image URL Textbox Label
        lblImageUrl = new Label() { Text = "Image Path:", Location = new System.Drawing.Point(10, 170), AutoSize = true }; // New label shifted down
        this.Controls.Add(lblImageUrl);

        // Image URL Textbox
        txtImageUrl = new TextBox()
        {
            Location = new System.Drawing.Point(10, 190),
            Size = new System.Drawing.Size(460, 20),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            BorderStyle = BorderStyle.FixedSingle
        };
        // Explicitly set colors for text boxes
        txtImageUrl.BackColor = _buttonActiveColor;
        txtImageUrl.ForeColor = _foreColor;
        this.Controls.Add(txtImageUrl);


        // Save Button
        btnSave = new Button()
        {
            Text = "Save",
            Location = new System.Drawing.Point(300, 240), // Shifted down
            Width = 80,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnSave.Click += BtnSave_Click;
        this.Controls.Add(btnSave);

        // Cancel Button
        btnCancel = new Button()
        {
            Text = "Cancel",
            Location = new System.Drawing.Point(390, 240), // Shifted down
            Width = 80,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnCancel.Click += BtnCancel_Click;

        this.Controls.Add(btnCancel);
    }

    private void ApplyTheme()
    {
        this.BackColor = _backColor;
        this.ForeColor = _foreColor;

        ApplyThemeToControls(this.Controls);
    }

    private void ApplyThemeToControls(Control.ControlCollection controls)
    {
        foreach (Control control in controls)
        {
            control.ForeColor = _foreColor; // Set ForeColor universally

            if (control is Button button)
            {
                // Unsubscribe to avoid double-subscription
                button.MouseEnter -= Button_MouseEnter;
                button.MouseLeave -= Button_MouseLeave;
                button.MouseDown -= Button_MouseDown;
                button.MouseUp -= Button_MouseUp;

                // Subscribe to events for hover/active effects
                button.MouseEnter += Button_MouseEnter;
                button.MouseLeave += Button_MouseLeave;
                button.MouseDown += Button_MouseDown;
                button.MouseUp += Button_MouseUp;

                button.FlatStyle = FlatStyle.Flat; // Use Flat style for custom border
                button.BackColor = _controlBackColor;
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 90); // Subtle border
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = _buttonActiveColor;
                textBox.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is NumericUpDown numericUpDownControl)
            {
                numericUpDownControl.BackColor = _buttonActiveColor;
                numericUpDownControl.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is Label label)
            {
                label.BackColor = Color.Transparent;
            }
            // No explicit RichTextBox for EditMessageForm, but adding for completeness if it were ever added.

            if (control.HasChildren)
            {
                ApplyThemeToControls(control.Controls);
            }
        }
    }

    // Button event handlers for hover/active effects specific to EditMessageForm
    private void Button_MouseEnter(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture)
        {
            button.BackColor = SettingsManager.GetButtonHoverColor();
        }
    }

    private void Button_MouseLeave(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture)
        {
            button.BackColor = _controlBackColor; // Use the form's control back color
        }
    }

    private void Button_MouseDown(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            button.BackColor = SettingsManager.GetButtonActiveColor();
        }
    }

    private void Button_MouseUp(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            if (button.ClientRectangle.Contains(button.PointToClient(Cursor.Position)))
            {
                button.BackColor = SettingsManager.GetButtonHoverColor();
            }
            else
            {
                button.BackColor = _controlBackColor;
            }
        }
    }


    private void LoadMessageData()
    {
        if (EditedMessage != null)
        {
            txtMessageText.Text = EditedMessage.Text;
            nudMessageWeight.Value = EditedMessage.Weight;
            txtImageUrl.Text = EditedMessage.ImageUrl; // Load ImageUrl
        }
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        string newText = txtMessageText.Text.Trim();
        int newWeight = (int)nudMessageWeight.Value;
        string newImageUrl = txtImageUrl.Text.Trim(); // Get new ImageUrl

        if (string.IsNullOrEmpty(newText) && string.IsNullOrEmpty(newImageUrl))
        {
            MessageBox.Show("Message text and Image URL cannot both be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Update the properties of the original Message object directly
        EditedMessage.Text = newText;
        EditedMessage.Weight = newWeight;
        EditedMessage.ImageUrl = newImageUrl; // Update ImageUrl

        this.DialogResult = DialogResult.OK; // Indicate success
        this.Close(); // Close the form
    }

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel; // Indicate cancellation
        this.Close(); // Close the form
    }
}

/// <summary>
/// A form for configuring application settings such as colors and hotkeys.
/// </summary>
public class SettingsForm : Form
{
    private Settings _tempSettings; // A temporary copy of settings for editing
    private TableLayoutPanel settingsLayout;

    // UI controls for colors (associated with TextBoxes, Panels (swatch), and TrackBars)
    private TextBox txtBackColorHex;
    private Panel pnlBackColorSwatch;
    private TrackBar trkBackColorRed;
    private TrackBar trkBackColorGreen;
    private TrackBar trkBackColorBlue;

    private TextBox txtForeColorHex;
    private Panel pnlForeColorSwatch;
    private TrackBar trkForeColorRed;
    private TrackBar trkForeColorGreen;
    private TrackBar trkForeColorBlue;

    private TextBox txtControlBackColorHex;
    private Panel pnlControlBackColorSwatch;
    private TrackBar trkControlBackColorRed;
    private TrackBar trkControlBackColorGreen;
    private TrackBar trkControlBackColorBlue;

    private TextBox txtButtonHoverColorHex;
    private Panel pnlButtonHoverColorSwatch;
    private TrackBar trkButtonHoverColorRed;
    private TrackBar trkButtonHoverColorGreen;
    private TrackBar trkButtonHoverColorBlue;

    private TextBox txtButtonActiveColorHex;
    private Panel pnlButtonActiveColorSwatch;
    private TrackBar trkButtonActiveColorRed;
    private TrackBar trkButtonActiveColorGreen;
    private TrackBar trkButtonActiveColorBlue;

    private TextBox txtDataGridViewHeaderBackColorHex;
    private Panel pnlDataGridViewHeaderBackColorSwatch;
    private TrackBar trkDataGridViewHeaderBackColorRed;
    private TrackBar trkDataGridViewHeaderBackColorGreen;
    private TrackBar trkDataGridViewHeaderBackColorBlue;

    private TextBox txtHighlightColorHex;
    private Panel pnlHighlightColorSwatch;
    private TrackBar trkHighlightColorRed;
    private TrackBar trkHighlightColorGreen;
    private TrackBar trkHighlightColorBlue;


    // UI controls for hotkeys
    private TextBox txtSendHotkey;
    private TextBox txtSaveHotkey;
    private TextBox txtLoadHotkey;
    private TextBox txtAddHotkey;
    private TextBox txtDuplicateHotkey;
    private TextBox txtDeleteHotkey;
    private TextBox txtEditHotkey;

    // UI control for Send Options
    private CheckBox chkAllowDuplicateMessages;

    private Button btnSave;
    private Button btnCancel;
    private Button btnResetToDefaults;


    public SettingsForm(Settings currentSettings)
    {
        // Create a deep copy of the settings to allow cancelling changes
        _tempSettings = JsonConvert.DeserializeObject<Settings>(JsonConvert.SerializeObject(currentSettings));
        InitializeComponent();
        LoadSettingsToUI();
        ApplyThemeToForm(); // Apply theme to settings form itself
    }

    private void InitializeComponent()
    {
        this.Text = "Application Settings";
        this.Size = new System.Drawing.Size(600, 850); // Increased height to prevent cutoff
        this.MinimumSize = new System.Drawing.Size(550, 750); // Adjusted minimum size
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.Fixed3D; // Made resizable
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        settingsLayout = new TableLayoutPanel();
        settingsLayout.Dock = DockStyle.Fill;
        settingsLayout.Padding = new Padding(15); // Increased padding for general layout
        settingsLayout.ColumnCount = 1;
        settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        settingsLayout.AutoScroll = true; // Make the entire layout scrollable
        settingsLayout.BackColor = SettingsManager.GetBackColor(); // Apply theme color to the main layout

        settingsLayout.RowCount = 0; // Start with 0, add rows dynamically

        // --- Color Settings Section ---
        AddSectionHeader("Color Settings");

        TableLayoutPanel pnlColorSettings = new TableLayoutPanel();
        pnlColorSettings.Dock = DockStyle.Top;
        pnlColorSettings.AutoSize = true;
        pnlColorSettings.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        pnlColorSettings.ColumnCount = 1;
        pnlColorSettings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        pnlColorSettings.BackColor = SettingsManager.GetControlBackColor(); // Apply control background color
        pnlColorSettings.Padding = new Padding(5); // Internal padding

        AddColorSetting("Background Color:", out txtBackColorHex, out pnlBackColorSwatch, out trkBackColorRed, out trkBackColorGreen, out trkBackColorBlue, (color) => _tempSettings.BackColorArgb = color.ToArgb(), _tempSettings.BackColorArgb, pnlColorSettings);
        AddColorSetting("Foreground Color:", out txtForeColorHex, out pnlForeColorSwatch, out trkForeColorRed, out trkForeColorGreen, out trkForeColorBlue, (color) => _tempSettings.ForeColorArgb = color.ToArgb(), _tempSettings.ForeColorArgb, pnlColorSettings);
        AddColorSetting("Control Back Color:", out txtControlBackColorHex, out pnlControlBackColorSwatch, out trkControlBackColorRed, out trkControlBackColorGreen, out trkControlBackColorBlue, (color) => _tempSettings.ControlBackColorArgb = color.ToArgb(), _tempSettings.ControlBackColorArgb, pnlColorSettings);
        AddColorSetting("Button Hover Color:", out txtButtonHoverColorHex, out pnlButtonHoverColorSwatch, out trkButtonHoverColorRed, out trkButtonHoverColorGreen, out trkButtonHoverColorBlue, (color) => _tempSettings.ButtonHoverColorArgb = color.ToArgb(), _tempSettings.ButtonHoverColorArgb, pnlColorSettings);
        AddColorSetting("Button Active Color:", out txtButtonActiveColorHex, out pnlButtonActiveColorSwatch, out trkButtonActiveColorRed, out trkButtonActiveColorGreen, out trkButtonActiveColorBlue, (color) => _tempSettings.ButtonActiveColorArgb = color.ToArgb(), _tempSettings.ButtonActiveColorArgb, pnlColorSettings);
        AddColorSetting("DGV Header Color:", out txtDataGridViewHeaderBackColorHex, out pnlDataGridViewHeaderBackColorSwatch, out trkDataGridViewHeaderBackColorRed, out trkDataGridViewHeaderBackColorGreen, out trkDataGridViewHeaderBackColorBlue, (color) => _tempSettings.DataGridViewHeaderBackColorArgb = color.ToArgb(), _tempSettings.DataGridViewHeaderBackColorArgb, pnlColorSettings);
        AddColorSetting("Highlight Color:", out txtHighlightColorHex, out pnlHighlightColorSwatch, out trkHighlightColorRed, out trkHighlightColorGreen, out trkHighlightColorBlue, (color) => _tempSettings.HighlightColorArgb = color.ToArgb(), _tempSettings.HighlightColorArgb, pnlColorSettings);

        settingsLayout.Controls.Add(pnlColorSettings, 0, settingsLayout.RowCount);
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        settingsLayout.RowCount++;


        // Add some vertical spacing to main settingsLayout
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        settingsLayout.RowCount++;

        // --- Hotkey Settings Section ---
        AddSectionHeader("Hotkey Settings (Enter Key Name, e.g., F5, S, L, Enter, Delete)");

        TableLayoutPanel pnlHotkeySettings = new TableLayoutPanel();
        pnlHotkeySettings.Dock = DockStyle.Top;
        pnlHotkeySettings.AutoSize = true;
        pnlHotkeySettings.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        pnlHotkeySettings.ColumnCount = 1;
        pnlHotkeySettings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        pnlHotkeySettings.BackColor = SettingsManager.GetControlBackColor(); // Apply control background color
        pnlHotkeySettings.Padding = new Padding(5);

        AddHotkeySetting("Send Messages (F5):", out txtSendHotkey, _tempSettings.SendHotkey, (text) => _tempSettings.SendHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Export Messages (Ctrl+S):", out txtSaveHotkey, _tempSettings.SaveHotkey, (text) => _tempSettings.SaveHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Import Messages (Ctrl+L):", out txtLoadHotkey, _tempSettings.LoadHotkey, (text) => _tempSettings.LoadHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Add Message (Ctrl+N):", out txtAddHotkey, _tempSettings.AddHotkey, (text) => _tempSettings.AddHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Duplicate Message (Ctrl+D):", out txtDuplicateHotkey, _tempSettings.DuplicateHotkey, (text) => _tempSettings.DuplicateHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Delete Message (Del):", out txtDeleteHotkey, _tempSettings.DeleteHotkey, (text) => _tempSettings.DeleteHotkey = text, pnlHotkeySettings);
        AddHotkeySetting("Edit Message (Enter on DGV):", out txtEditHotkey, _tempSettings.EditHotkey, (text) => _tempSettings.EditHotkey = text, pnlHotkeySettings);

        settingsLayout.Controls.Add(pnlHotkeySettings, 0, settingsLayout.RowCount);
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        settingsLayout.RowCount++;

        // Add some vertical spacing to main settingsLayout
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        settingsLayout.RowCount++;

        // --- Send Options Section ---
        AddSectionHeader("Send Options");

        TableLayoutPanel pnlSendOptions = new TableLayoutPanel();
        pnlSendOptions.Dock = DockStyle.Top;
        pnlSendOptions.AutoSize = true;
        pnlSendOptions.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        pnlSendOptions.ColumnCount = 1;
        pnlSendOptions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        pnlSendOptions.BackColor = SettingsManager.GetControlBackColor();
        pnlSendOptions.Padding = new Padding(5);

        // Allow Duplicate Messages Checkbox
        chkAllowDuplicateMessages = new CheckBox()
        {
            Text = "Allow sending the same message multiple times (do not remove from pool)",
            Checked = _tempSettings.AllowDuplicateMessages,
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };
        chkAllowDuplicateMessages.CheckedChanged += (sender, e) =>
        {
            _tempSettings.AllowDuplicateMessages = chkAllowDuplicateMessages.Checked;
        };
        pnlSendOptions.Controls.Add(chkAllowDuplicateMessages, 0, 0);
        pnlSendOptions.RowStyles.Add(new RowStyle(SizeType.AutoSize));


        settingsLayout.Controls.Add(pnlSendOptions, 0, settingsLayout.RowCount);
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        settingsLayout.RowCount++;


        // Add some vertical spacing before buttons
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        settingsLayout.RowCount++;

        // --- Control Buttons ---
        TableLayoutPanel buttonPanel = new TableLayoutPanel();
        buttonPanel.Dock = DockStyle.Fill;
        buttonPanel.ColumnCount = 3;
        buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
        buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
        buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
        buttonPanel.RowCount = 1;
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));

        btnSave = new Button() { Text = "Save Settings", Dock = DockStyle.Fill };
        btnSave.Click += BtnSave_Click;
        buttonPanel.Controls.Add(btnSave, 0, 0);

        btnCancel = new Button() { Text = "Cancel", Dock = DockStyle.Fill };
        btnCancel.Click += BtnCancel_Click;
        buttonPanel.Controls.Add(btnCancel, 1, 0);

        btnResetToDefaults = new Button() { Text = "Reset to Defaults", Dock = DockStyle.Fill };
        btnResetToDefaults.Click += BtnResetToDefaults_Click;
        buttonPanel.Controls.Add(btnResetToDefaults, 2, 0);

        settingsLayout.Controls.Add(buttonPanel, 0, settingsLayout.RowCount);
        settingsLayout.SetColumnSpan(buttonPanel, 1);
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
        settingsLayout.RowCount++;


        this.Controls.Add(settingsLayout);
    }

    private void AddSectionHeader(string text)
    {
        // Add a vertical spacer before the header
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10F));
        settingsLayout.RowCount++;

        Label header = new Label()
        {
            Text = text,
            Font = new Font(this.Font, FontStyle.Bold),
            AutoSize = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Padding = new Padding(0, 5, 0, 0) // Padding at top
        };
        settingsLayout.Controls.Add(header, 0, settingsLayout.RowCount);
        settingsLayout.SetColumnSpan(header, 1);
        settingsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Auto-size to content for header
        settingsLayout.RowCount++;
    }

    /// <summary>
    /// Adds controls for a color setting including a label, hex textbox, color swatch, and RGB sliders.
    /// </summary>
    /// <param name="labelText">The text for the color setting label.</param>
    /// <param name="hexTextBoxRef">Reference to the created hex textbox.</param>
    /// <param name="colorSwatchRef">Reference to the created color swatch panel.</param>
    /// <param name="redSliderRef">Reference to the created Red component slider.</param>
    /// <param name="greenSliderRef">Reference to the created Green component slider.</param>
    /// <param name="blueSliderRef">Reference to the created Blue component slider.</param>
    /// <param name="setColorAction">Action to update the settings object with the new color.</param>
    /// <param name="initialColorArgb">Initial ARGB value for the color.</param>
    /// <param name="parentLayout">The TableLayoutPanel to which these color controls should be added.</param>
    private void AddColorSetting(string labelText, out TextBox hexTextBoxRef, out Panel colorSwatchRef, out TrackBar redSliderRef, out TrackBar greenSliderRef, out TrackBar blueSliderRef, Action<Color> setColorAction, int initialColorArgb, TableLayoutPanel parentLayout)
    {
        // Create a nested TableLayoutPanel to hold ALL controls for THIS SINGLE color setting
        TableLayoutPanel colorSettingGroupPanel = new TableLayoutPanel();
        colorSettingGroupPanel.Dock = DockStyle.Top; // Dock to top within parentLayout
        colorSettingGroupPanel.AutoSize = true;      // Let it size itself vertically
        colorSettingGroupPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink; // Important for efficient sizing
        colorSettingGroupPanel.ColumnCount = 2;
        colorSettingGroupPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60F)); // For R, G, B labels
        colorSettingGroupPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // For main controls
        colorSettingGroupPanel.BackColor = Color.Transparent; // Inherit background from parent
        colorSettingGroupPanel.Padding = new Padding(0, 5, 0, 5); // Small vertical padding between color groups

        // Row 0: The main label for this color setting
        Label mainLabel = new Label()
        {
            Text = labelText,
            Font = new Font(this.Font, FontStyle.Bold), // Make the main label bold
            Anchor = AnchorStyles.Left | AnchorStyles.Bottom, // Anchor to bottom to push it down slightly if row is tall
            AutoSize = true
        };
        colorSettingGroupPanel.Controls.Add(mainLabel, 0, 0);
        colorSettingGroupPanel.SetColumnSpan(mainLabel, 2); // Span both columns
        colorSettingGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 25F)); // Fixed height for the label row

        // Row 1: Hex Textbox and Color Swatch (in a sub-panel for layout flexibility)
        TableLayoutPanel hexSwatchPanel = new TableLayoutPanel();
        hexSwatchPanel.Dock = DockStyle.Fill;
        hexSwatchPanel.ColumnCount = 2;
        hexSwatchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F)); // Hex textbox
        hexSwatchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F)); // Swatch
        hexSwatchPanel.RowCount = 1;
        hexSwatchPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        hexSwatchPanel.BackColor = Color.Transparent;

        TextBox hexTextBox = new TextBox()
        {
            Text = ColorToHex(Color.FromArgb(initialColorArgb)),
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = SettingsManager.GetButtonActiveColor(),
            ForeColor = SettingsManager.GetForeColor()
        };
        hexTextBoxRef = hexTextBox;
        hexSwatchPanel.Controls.Add(hexTextBox, 0, 0);

        Panel colorSwatch = new Panel()
        {
            BackColor = Color.FromArgb(initialColorArgb),
            BorderStyle = BorderStyle.FixedSingle,
            Dock = DockStyle.Fill,
            Margin = new Padding(3),
        };
        colorSwatchRef = colorSwatch;
        hexSwatchPanel.Controls.Add(colorSwatch, 1, 0);

        colorSettingGroupPanel.Controls.Add(hexSwatchPanel, 0, 1); // Add hex/swatch panel
        colorSettingGroupPanel.SetColumnSpan(hexSwatchPanel, 2); // Span both columns
        colorSettingGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F)); // Fixed height for hex/swatch row


        // Row 2: Red Slider
        Label lblRed = new Label() { Text = "R:", Anchor = AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom, TextAlign = ContentAlignment.MiddleRight };
        colorSettingGroupPanel.Controls.Add(lblRed, 0, 2);
        TrackBar redSlider = new TrackBar()
        {
            Minimum = 0,
            Maximum = 255,
            Value = Color.FromArgb(initialColorArgb).R,
            TickFrequency = 32,
            LargeChange = 16,
            SmallChange = 1,
            Dock = DockStyle.Fill,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BackColor = SettingsManager.GetControlBackColor(),
            ForeColor = SettingsManager.GetForeColor(),
        };
        redSliderRef = redSlider;
        colorSettingGroupPanel.Controls.Add(redSlider, 1, 2);
        colorSettingGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F)); // Fixed height for slider row


        // Row 3: Green Slider
        Label lblGreen = new Label() { Text = "G:", Anchor = AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom, TextAlign = ContentAlignment.MiddleRight };
        colorSettingGroupPanel.Controls.Add(lblGreen, 0, 3);
        TrackBar greenSlider = new TrackBar()
        {
            Minimum = 0,
            Maximum = 255,
            Value = Color.FromArgb(initialColorArgb).G,
            TickFrequency = 32,
            LargeChange = 16,
            SmallChange = 1,
            Dock = DockStyle.Fill,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BackColor = SettingsManager.GetControlBackColor(),
            ForeColor = SettingsManager.GetForeColor(),
        };
        greenSliderRef = greenSlider;
        colorSettingGroupPanel.Controls.Add(greenSlider, 1, 3);
        colorSettingGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F)); // Fixed height for slider row


        // Row 4: Blue Slider
        Label lblBlue = new Label() { Text = "B:", Anchor = AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom, TextAlign = ContentAlignment.MiddleRight };
        colorSettingGroupPanel.Controls.Add(lblBlue, 0, 4);
        TrackBar blueSlider = new TrackBar()
        {
            Minimum = 0,
            Maximum = 255,
            Value = Color.FromArgb(initialColorArgb).B,
            TickFrequency = 32,
            LargeChange = 16,
            SmallChange = 1,
            Dock = DockStyle.Fill,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BackColor = SettingsManager.GetControlBackColor(),
            ForeColor = SettingsManager.GetForeColor(),
        };
        blueSliderRef = blueSlider;
        colorSettingGroupPanel.Controls.Add(blueSlider, 1, 4);
        colorSettingGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F)); // Fixed height for slider row


        // NEW: Flag to prevent re-entrancy for this specific color setting's controls
        bool ignoreColorChange = false;

        // Hex Textbox TextChanged event handler
        hexTextBox.TextChanged += (sender, e) =>
        {
            if (ignoreColorChange) return;

            ignoreColorChange = true;

            Color newColor;
            if (HexToColor(hexTextBox.Text, out newColor))
            {
                colorSwatch.BackColor = newColor;
                redSlider.Value = newColor.R;
                greenSlider.Value = newColor.G;
                blueSlider.Value = newColor.B;
                setColorAction(newColor);
                hexTextBox.BackColor = SettingsManager.GetButtonActiveColor();
            }
            else
            {
                hexTextBox.BackColor = Color.IndianRed;
            }

            ignoreColorChange = false;
        };

        // Slider Scroll event handler (reusable for R, G, B)
        EventHandler sliderScrollHandler = (sender, e) =>
        {
            if (ignoreColorChange) return;

            ignoreColorChange = true;

            Color newColor = Color.FromArgb(255, redSlider.Value, greenSlider.Value, blueSlider.Value);
            hexTextBox.Text = ColorToHex(newColor);
            colorSwatch.BackColor = newColor;
            setColorAction(newColor);

            ignoreColorChange = false;
        };

        redSlider.Scroll += sliderScrollHandler;
        greenSlider.Scroll += sliderScrollHandler;
        blueSlider.Scroll += sliderScrollHandler;

        // Add the entire color setting group panel to the parent layout (pnlColorSettings)
        parentLayout.Controls.Add(colorSettingGroupPanel, 0, parentLayout.RowCount);
        parentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // This row will auto-size to fit colorSettingGroupPanel
        parentLayout.RowCount++;
    }


    /// <summary>
    /// Adds controls for a hotkey setting including a label and a textbox.
    /// </summary>
    /// <param name="labelText">The text for the hotkey setting label.</param>
    /// <param name="hotkeyTextBoxRef">Reference to the created hotkey textbox.</param>
    /// <param name="initialHotkey">Initial hotkey string.</param>
    /// <param name="setHotkeyAction">Action to update the settings object with the new hotkey.</param>
    /// <param name="parentLayout">The TableLayoutPanel to which these hotkey controls should be added.</param>
    private void AddHotkeySetting(string labelText, out TextBox hotkeyTextBoxRef, string initialHotkey, Action<string> setHotkeyAction, TableLayoutPanel parentLayout)
    {
        TableLayoutPanel hotkeyGroupPanel = new TableLayoutPanel();
        hotkeyGroupPanel.Dock = DockStyle.Top;
        hotkeyGroupPanel.AutoSize = true;
        hotkeyGroupPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        hotkeyGroupPanel.ColumnCount = 2;
        hotkeyGroupPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Label column
        hotkeyGroupPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Textbox column
        hotkeyGroupPanel.BackColor = Color.Transparent; // Inherit from parent
        hotkeyGroupPanel.Padding = new Padding(0, 3, 0, 3); // Small padding

        Label label = new Label() { Text = labelText, Anchor = AnchorStyles.Left };
        hotkeyGroupPanel.Controls.Add(label, 0, 0);

        TextBox hotkeyTextBox = new TextBox()
        {
            Text = initialHotkey,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign = HorizontalAlignment.Center,
            BackColor = SettingsManager.GetButtonActiveColor(),
            ForeColor = SettingsManager.GetForeColor()
        };
        hotkeyTextBoxRef = hotkeyTextBox;

        hotkeyTextBox.KeyDown += (sender, e) =>
        {
            e.SuppressKeyPress = true; // Prevent the key from being processed by the system
            e.Handled = true; // Mark the event as handled

            // Determine the key combination string
            string keyName = e.KeyCode.ToString();

            // Store the key name directly, assuming single keys or combinations handled by KeyEventArgs
            hotkeyTextBox.Text = keyName;
            setHotkeyAction(keyName); // Update the temporary settings object
        };

        hotkeyGroupPanel.Controls.Add(hotkeyTextBox, 1, 0);
        hotkeyGroupPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F)); // Fixed height for hotkey row

        parentLayout.Controls.Add(hotkeyGroupPanel, 0, parentLayout.RowCount);
        parentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        parentLayout.RowCount++;
    }

    // Helper to convert Color to Hex string
    private string ColorToHex(Color color)
    {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    // Helper to convert Hex string to Color
    private bool HexToColor(string hex, out Color color)
    {
        color = Color.Black; // Default value
        if (string.IsNullOrWhiteSpace(hex) || (hex.Length != 7 && hex.Length != 9))
            return false;

        // Remove '#' if present
        if (hex.StartsWith("#"))
            hex = hex.Substring(1);

        try
        {
            int r = 0, g = 0, b = 0, a = 255; // Default alpha to 255 (fully opaque)

            if (hex.Length == 6) // RRGGBB
            {
                r = int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                g = int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                b = int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
            }
            else if (hex.Length == 8) // AARRGGBB or RRGGBBAA (usually AARRGGBB)
            {
                a = int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                r = int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                g = int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                b = int.Parse(hex.Substring(6, 2), System.Globalization.NumberStyles.HexNumber);
            }
            else
            {
                return false;
            }

            color = Color.FromArgb(a, r, g, b);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private Color GetContrastColor(Color color)
    {
        // Calculate the perceptive luminance (0-255)
        int L = (int)(0.299 * color.R + 0.587 * color.G + 0.114 * color.B);
        return L > 186 ? Color.Black : Color.White; // Use a threshold for light/dark contrast
    }

    private void LoadSettingsToUI()
    {
        // Load colors (now updating hex and all three sliders)
        Color backColor = Color.FromArgb(_tempSettings.BackColorArgb);
        txtBackColorHex.Text = ColorToHex(backColor);
        pnlBackColorSwatch.BackColor = backColor;
        trkBackColorRed.Value = backColor.R;
        trkBackColorGreen.Value = backColor.G;
        trkBackColorBlue.Value = backColor.B;

        Color foreColor = Color.FromArgb(_tempSettings.ForeColorArgb);
        txtForeColorHex.Text = ColorToHex(foreColor);
        pnlForeColorSwatch.BackColor = foreColor;
        trkForeColorRed.Value = foreColor.R;
        trkForeColorGreen.Value = foreColor.G;
        trkForeColorBlue.Value = foreColor.B;

        Color controlBackColor = Color.FromArgb(_tempSettings.ControlBackColorArgb);
        txtControlBackColorHex.Text = ColorToHex(controlBackColor);
        pnlControlBackColorSwatch.BackColor = controlBackColor;
        trkControlBackColorRed.Value = controlBackColor.R;
        trkControlBackColorGreen.Value = controlBackColor.G;
        trkControlBackColorBlue.Value = controlBackColor.B;

        Color buttonHoverColor = Color.FromArgb(_tempSettings.ButtonHoverColorArgb);
        txtButtonHoverColorHex.Text = ColorToHex(buttonHoverColor);
        pnlButtonHoverColorSwatch.BackColor = buttonHoverColor;
        trkButtonHoverColorRed.Value = buttonHoverColor.R;
        trkButtonHoverColorGreen.Value = buttonHoverColor.G;
        trkButtonHoverColorBlue.Value = buttonHoverColor.B;

        Color buttonActiveColor = Color.FromArgb(_tempSettings.ButtonActiveColorArgb);
        txtButtonActiveColorHex.Text = ColorToHex(buttonActiveColor);
        pnlButtonActiveColorSwatch.BackColor = buttonActiveColor;
        trkButtonActiveColorRed.Value = buttonActiveColor.R;
        trkButtonActiveColorGreen.Value = buttonActiveColor.G;
        trkButtonActiveColorBlue.Value = buttonActiveColor.B;

        Color dataGridViewHeaderBackColor = Color.FromArgb(_tempSettings.DataGridViewHeaderBackColorArgb);
        txtDataGridViewHeaderBackColorHex.Text = ColorToHex(dataGridViewHeaderBackColor);
        pnlDataGridViewHeaderBackColorSwatch.BackColor = dataGridViewHeaderBackColor;
        trkDataGridViewHeaderBackColorRed.Value = dataGridViewHeaderBackColor.R;
        trkDataGridViewHeaderBackColorGreen.Value = dataGridViewHeaderBackColor.G;
        trkDataGridViewHeaderBackColorBlue.Value = dataGridViewHeaderBackColor.B;

        Color highlightColor = Color.FromArgb(_tempSettings.HighlightColorArgb);
        txtHighlightColorHex.Text = ColorToHex(highlightColor);
        pnlHighlightColorSwatch.BackColor = highlightColor;
        trkHighlightColorRed.Value = highlightColor.R;
        trkHighlightColorGreen.Value = highlightColor.G;
        trkHighlightColorBlue.Value = highlightColor.B;

        // Load hotkeys
        txtSendHotkey.Text = _tempSettings.SendHotkey;
        txtSaveHotkey.Text = _tempSettings.SaveHotkey;
        txtLoadHotkey.Text = _tempSettings.LoadHotkey;
        txtAddHotkey.Text = _tempSettings.AddHotkey;
        txtDuplicateHotkey.Text = _tempSettings.DuplicateHotkey;
        txtDeleteHotkey.Text = _tempSettings.DeleteHotkey;
        txtEditHotkey.Text = _tempSettings.EditHotkey;

        // Load Send Options
        chkAllowDuplicateMessages.Checked = _tempSettings.AllowDuplicateMessages;
    }

    private void ApplyThemeToForm()
    {
        this.BackColor = SettingsManager.GetBackColor();
        this.ForeColor = SettingsManager.GetForeColor();

        // Recursively apply theme to controls, ensuring appropriate colors
        ApplyThemeToControls(this.Controls);
    }

    private void ApplyThemeToControls(Control.ControlCollection controls)
    {
        foreach (Control control in controls)
        {
            control.ForeColor = SettingsManager.GetForeColor();

            if (control is Button button)
            {
                // Unsubscribe to avoid double-subscription
                button.MouseEnter -= Button_MouseEnter;
                button.MouseLeave -= Button_MouseLeave;
                button.MouseDown -= Button_MouseDown;
                button.MouseUp -= Button_MouseUp;

                // Subscribe to events for hover/active effects
                button.MouseEnter += Button_MouseEnter;
                button.MouseLeave += Button_MouseLeave;
                button.MouseDown += Button_MouseDown;
                button.MouseUp += Button_MouseUp;

                // Buttons in settings form will use their own backcolor from tempSettings,
                // but other buttons (Save, Cancel, Reset) will use control back color.
                if (button == btnSave || button == btnCancel || button == btnResetToDefaults)
                {
                    button.BackColor = SettingsManager.GetControlBackColor();
                }
                button.FlatStyle = FlatStyle.Flat; // Ensure flat style for borders
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 90); // Subtle border
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = SettingsManager.GetButtonActiveColor();
                textBox.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is NumericUpDown numericUpDown)
            {
                numericUpDown.BackColor = SettingsManager.GetButtonActiveColor();
                numericUpDown.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is Label label)
            {
                label.BackColor = Color.Transparent;
            }
            else if (control is Panel || control is TableLayoutPanel)
            {
                control.BackColor = SettingsManager.GetControlBackColor();
            }
            else if (control is TrackBar trackBar) // New: Apply theme to TrackBar
            {
                trackBar.BackColor = SettingsManager.GetControlBackColor();
                trackBar.ForeColor = SettingsManager.GetForeColor();
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.BackColor = Color.Transparent;
            }

            if (control.HasChildren)
            {
                ApplyThemeToControls(control.Controls);
            }
        }
    }

    // Button event handlers for hover/active effects specific to SettingsForm
    private void Button_MouseEnter(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture)
        {
            if (button == btnSave || button == btnCancel || button == btnResetToDefaults)
            {
                button.BackColor = SettingsManager.GetButtonHoverColor();
            }
            // For color swatch panels, their backcolor represents the color, so no hover effect.
        }
    }

    private void Button_MouseLeave(object sender, EventArgs e)
    {
        Button button = sender as Button;
        if (button != null && !button.Capture)
        {
            if (button == btnSave || button == btnCancel || button == btnResetToDefaults)
            {
                button.BackColor = SettingsManager.GetControlBackColor();
            }
        }
    }

    private void Button_MouseDown(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            if (button == btnSave || button == btnCancel || button == btnResetToDefaults)
            {
                button.BackColor = SettingsManager.GetButtonActiveColor();
            }
        }
    }

    private void Button_MouseUp(object sender, MouseEventArgs e)
    {
        Button button = sender as Button;
        if (button != null && e.Button == MouseButtons.Left)
        {
            if (button.ClientRectangle.Contains(button.PointToClient(Cursor.Position)))
            {
                button.BackColor = SettingsManager.GetButtonHoverColor();
            }
            else
            {
                button.BackColor = SettingsManager.GetControlBackColor();
            }
        }
    }


    private void BtnSave_Click(object sender, EventArgs e)
    {
        // Hotkey validation (basic check for empty strings)
        if (string.IsNullOrWhiteSpace(_tempSettings.SendHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.SaveHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.LoadHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.AddHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.DuplicateHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.DeleteHotkey) ||
            string.IsNullOrWhiteSpace(_tempSettings.EditHotkey))
        {
            MessageBox.Show("All hotkeys must have a value. Please ensure no hotkey field is empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Attempt to parse all hotkeys to ensure they are valid Keys enum values
        foreach (var prop in _tempSettings.GetType().GetProperties())
        {
            if (prop.Name.EndsWith("Hotkey") && prop.PropertyType == typeof(string))
            {
                string hotkeyString = (string)prop.GetValue(_tempSettings);
                if (!Enum.IsDefined(typeof(Keys), hotkeyString))
                {
                    MessageBox.Show($"Invalid hotkey value for '{prop.Name}': '{hotkeyString}'. Please enter a valid System.Windows.Forms.Keys enum member name (e.g., 'F5', 'A', 'Enter', 'Delete').", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        // Save the temporary settings
        SettingsManager.SaveSettings(_tempSettings);
        this.DialogResult = DialogResult.OK;
        this.Close();
    }

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel;
        this.Close();
    }

    private void BtnResetToDefaults_Click(object sender, EventArgs e)
    {
        DialogResult confirmResult = MessageBox.Show(
            "Are you sure you want to reset all settings to their default values?",
            "Confirm Reset",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirmResult == DialogResult.Yes)
        {
            _tempSettings = new Settings(); // Create a new Settings object with default values
            LoadSettingsToUI(); // Update the UI with the default values
            MessageBox.Show("Settings have been reset to defaults. Click 'Save Settings' to apply.", "Reset Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
