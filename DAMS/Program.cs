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
using System.Threading.Tasks; // For using Task-based asynchronous operations
using System.IO; // For file operations
using Newtonsoft.Json; // For JSON serialization/deserialization.
using System.Drawing; // For Color

// Represents a message with associated text and a weight for rarity selection.
// Moved outside MainForm to be accessible by EditMessageForm.
public class Message
{
    public string Text { get; set; }
    public int Weight { get; set; }

    public Message(string text, int weight)
    {
        Text = text;
        Weight = weight;
    }
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
    private Button btnAdd;
    private Button btnUpdate;
    private Button btnDelete;
    private Button btnStartStop;
    private Label lblStatus;
    private Button btnSaveMessages; // New Save button for export
    private Button btnLoadMessages; // New Load button for import
    private NumericUpDown nudLoopCount; // New control for loop count
    private Label lblLoopCount; // Label for loop count
    private Label lblText; // Declared as a field to fix CS0103
    private Label lblWeight; // Declared as a field to fix CS0103


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

    // Constants for delays
    const int FocusDelayMs = 1000; // Delay after focusing Discord window
    const int InputDelayMs = 500;  // Delay after typing each line or sending command

    // Dark Mode Colors
    private static readonly Color DarkBackColor = Color.FromArgb(44, 47, 51); // Discord-like dark grey
    private static readonly Color DarkForeColor = Color.FromArgb(220, 221, 222); // Light text for contrast
    private static readonly Color DarkControlBackColor = Color.FromArgb(54, 57, 63); // Slightly lighter dark grey for controls
    private static readonly Color DarkButtonHoverColor = Color.FromArgb(64, 67, 73); // Button hover effect
    private static readonly Color DarkButtonActiveColor = Color.FromArgb(32, 35, 39); // Button active effect
    private static readonly Color DataGridViewHeaderBackColor = Color.FromArgb(32, 34, 37); // Darker header

    public MainForm()
    {
        InitializeComponent();
        ApplyTheme(); // Apply dark theme after initializing components
        InitializeData();
    }

    private void InitializeComponent()
    {
        this.Text = "Discord Message Sender";
        this.Size = new System.Drawing.Size(800, 600);
        this.MinimumSize = new System.Drawing.Size(700, 500); // Allow some resizing
        this.StartPosition = FormStartPosition.CenterScreen;

        // DataGridView for messages
        dgvMessages = new DataGridView();
        dgvMessages.Dock = DockStyle.Fill;
        dgvMessages.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dgvMessages.AllowUserToAddRows = false; // Prevent direct adding via DGV
        dgvMessages.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvMessages.MultiSelect = false;
        dgvMessages.ReadOnly = true; // Make it read-only for direct editing, use the separate edit window
        dgvMessages.SelectionChanged += DgvMessages_SelectionChanged;

        // DataGridView styling for dark mode
        dgvMessages.BackgroundColor = DarkBackColor;
        dgvMessages.GridColor = DarkControlBackColor;
        dgvMessages.DefaultCellStyle.BackColor = DarkControlBackColor;
        dgvMessages.DefaultCellStyle.ForeColor = DarkForeColor;
        dgvMessages.DefaultCellStyle.SelectionBackColor = Color.FromArgb(70, 73, 79); // A subtle selection color
        dgvMessages.DefaultCellStyle.SelectionForeColor = DarkForeColor;
        dgvMessages.ColumnHeadersDefaultCellStyle.BackColor = DataGridViewHeaderBackColor;
        dgvMessages.ColumnHeadersDefaultCellStyle.ForeColor = DarkForeColor;
        dgvMessages.EnableHeadersVisualStyles = false; // Required to apply custom header styles
        dgvMessages.RowHeadersDefaultCellStyle.BackColor = DataGridViewHeaderBackColor;
        dgvMessages.RowHeadersDefaultCellStyle.ForeColor = DarkForeColor;


        // Main control panel using TableLayoutPanel for better layout management
        TableLayoutPanel controlPanel = new TableLayoutPanel();
        controlPanel.Dock = DockStyle.Bottom;
        controlPanel.Height = 190; // Adjusted height for more controls (including loop count)
        controlPanel.Padding = new Padding(10);
        controlPanel.ColumnCount = 3; // Two columns for text/weight, one for buttons
        controlPanel.RowCount = 6; // Adjusted rows for loop count

        // Configure column and row styles for responsive layout
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Left column for labels
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // Middle column for input fields
        controlPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F)); // Right column for buttons

        for (int i = 0; i < controlPanel.RowCount; i++)
        {
            controlPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Fixed height for rows
        }
        controlPanel.RowStyles[1].SizeType = SizeType.AutoSize; // Allow message text row to auto-size
        controlPanel.RowStyles[4].SizeType = SizeType.Absolute; // Ensure status row has fixed height


        // Controls for adding new messages (New layout using TableLayoutPanel)
        lblText = new Label() { Text = "New Message Text:", Anchor = AnchorStyles.Left }; // Now uses the field
        controlPanel.Controls.Add(lblText, 0, 0); // Column 0, Row 0
        txtMessageText = new TextBox() { Multiline = true, ScrollBars = ScrollBars.Vertical, Height = 60, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom };
        controlPanel.Controls.Add(txtMessageText, 1, 0); // Column 1, Row 0
        controlPanel.SetRowSpan(txtMessageText, 2); // Span two rows for larger text box

        lblWeight = new Label() { Text = "New Message Weight:", Anchor = AnchorStyles.Left }; // Now uses the field
        controlPanel.Controls.Add(lblWeight, 0, 2); // Column 0, Row 2
        nudMessageWeight = new NumericUpDown() { Minimum = 1, Maximum = int.MaxValue, Value = 1000, Anchor = AnchorStyles.Left }; // Max value set to int.MaxValue
        controlPanel.Controls.Add(nudMessageWeight, 1, 2); // Column 1, Row 2

        // New Loop Count controls
        lblLoopCount = new Label() { Text = "Send Count:", Anchor = AnchorStyles.Left };
        controlPanel.Controls.Add(lblLoopCount, 0, 3); // Column 0, Row 3
        nudLoopCount = new NumericUpDown() { Minimum = 1, Maximum = 99999, Value = 1, Anchor = AnchorStyles.Left }; // Set a reasonable max value
        controlPanel.Controls.Add(nudLoopCount, 1, 3); // Column 1, Row 3


        // Buttons
        btnAdd = new Button() { Text = "Add New Message", Dock = DockStyle.Fill };
        btnAdd.Click += BtnAdd_Click;
        controlPanel.Controls.Add(btnAdd, 2, 0); // Column 2, Row 0

        btnUpdate = new Button() { Text = "Edit Selected Message", Dock = DockStyle.Fill, Enabled = false }; // Initially disabled
        btnUpdate.Click += BtnUpdate_Click;
        controlPanel.Controls.Add(btnUpdate, 2, 1); // Column 2, Row 1

        btnDelete = new Button() { Text = "Delete Selected Message", Dock = DockStyle.Fill, Enabled = false }; // Initially disabled
        btnDelete.Click += BtnDelete_Click;
        controlPanel.Controls.Add(btnDelete, 2, 2); // Column 2, Row 2

        btnStartStop = new Button() { Text = "Start Sending", Dock = DockStyle.Fill };
        btnStartStop.Click += BtnStartStop_Click;
        controlPanel.Controls.Add(btnStartStop, 2, 3); // Column 2, Row 3

        btnSaveMessages = new Button() { Text = "Export Messages", Dock = DockStyle.Fill }; // Changed text to Export
        btnSaveMessages.Click += BtnSaveMessages_Click;
        controlPanel.Controls.Add(btnSaveMessages, 0, 4); // Column 0, Row 4

        btnLoadMessages = new Button() { Text = "Import Messages", Dock = DockStyle.Fill }; // Changed text to Import
        btnLoadMessages.Click += BtnLoadMessages_Click;
        controlPanel.Controls.Add(btnLoadMessages, 1, 4); // Column 1, Row 4

        // Status Label
        lblStatus = new Label() { Text = "Status: Idle", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        controlPanel.Controls.Add(lblStatus, 0, 5); // Column 0, Row 5
        controlPanel.SetColumnSpan(lblStatus, 3); // Span across all columns

        this.Controls.Add(dgvMessages);
        this.Controls.Add(controlPanel);
    }

    private void ApplyTheme()
    {
        this.BackColor = DarkBackColor;
        this.ForeColor = DarkForeColor;

        // Apply theme to TableLayoutPanel
        // controlPanel's BackColor is automatically inherited from its parent unless set explicitly
        // controlPanel.BackColor = DarkControlBackColor; // Example if you want a different shade for the panel

        // Apply theme to all controls within the main form's Controls collection
        // This is a more robust way than setting individual properties for each control
        // as it applies to controls added later as well (if re-applied).
        ApplyThemeToControls(this.Controls);
    }

    private void ApplyThemeToControls(Control.ControlCollection controls)
    {
        foreach (Control control in controls)
        {
            control.BackColor = DarkControlBackColor;
            control.ForeColor = DarkForeColor;

            // Specific styling for certain control types
            if (control is Button button)
            {
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.BackColor = DarkControlBackColor; // Default button color
                button.ForeColor = DarkForeColor;
                // You could add MouseEnter/MouseLeave events for hover effects here if desired
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = DarkButtonActiveColor; // Darker background for input fields
                textBox.ForeColor = DarkForeColor;
            }
            else if (control is NumericUpDown numericUpDown)
            {
                numericUpDown.BackColor = DarkButtonActiveColor; // Darker background for input fields
                numericUpDown.ForeColor = DarkForeColor;
            }
            else if (control is Label label)
            {
                // Labels already have ForeColor set, but ensuring consistency
                label.ForeColor = DarkForeColor;
                label.BackColor = Color.Transparent; // Labels typically don't have a background
            }
            else if (control is DataGridView dgv)
            {
                // DataGridView styling is done directly in InitializeComponent for more granular control
                // No need to re-apply here, but including for completeness
            }

            // Recursively apply theme to child controls (e.g., controls within a TableLayoutPanel)
            if (control.HasChildren)
            {
                ApplyThemeToControls(control.Controls);
            }
        }
    }


    private void InitializeData()
    {
        // Define default messages. Only the ":3" message is hardcoded now.
        // Other messages are expected to be imported or added via the UI.
        allMessages = new List<Message>
        {
            new Message(":3", 6000)
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

        if (string.IsNullOrEmpty(text))
        {
            MessageBox.Show("Message text cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Message newMessage = new Message(text, weight);
        allMessages.Add(newMessage);
        messageBindingSource.ResetBindings(false); // Refresh DataGridView
        availableMessages = new List<Message>(allMessages); // Reset available messages when adding new
        txtMessageText.Clear();
        nudMessageWeight.Value = 1000;
        // No automatic save here, rely on explicit Save/Load
    }

    // Event handler for updating a selected message
    private void BtnUpdate_Click(object sender, EventArgs e)
    {
        if (dgvMessages.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a message to update.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Get the selected message object from the DataGridView's data source
        Message selectedMessage = (Message)dgvMessages.SelectedRows[0].DataBoundItem;

        // Create and show the new EditMessageForm
        using (EditMessageForm editForm = new EditMessageForm(selectedMessage))
        {
            // Pass current theme colors to the edit form
            editForm.SetThemeColors(DarkBackColor, DarkForeColor, DarkControlBackColor, DarkButtonActiveColor, DataGridViewHeaderBackColor);

            // If the user clicks 'Save' in the edit form
            if (editForm.ShowDialog() == DialogResult.OK)
            {
                // The 'selectedMessage' object itself has been updated by EditMessageForm
                // So we just need to refresh the DataGridView and reset available messages.
                messageBindingSource.ResetBindings(false); // Refresh DataGridView to reflect changes
                availableMessages = new List<Message>(allMessages); // Reset available messages when updating
                MessageBox.Show("Message updated successfully.", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // No automatic save here, rely on explicit Save/Load
            }
        }
    }

    // Event handler for deleting a selected message
    private void BtnDelete_Click(object sender, EventArgs e)
    {
        if (dgvMessages.SelectedRows.Count == 0)
        {
            MessageBox.Show("Please select a message to delete.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Confirmation dialog for deletion
        DialogResult confirmResult = MessageBox.Show(
            "Are you sure you want to delete the selected message?",
            "Confirm Delete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirmResult == DialogResult.Yes)
        {
            Message selectedMessage = (Message)dgvMessages.SelectedRows[0].DataBoundItem;
            allMessages.Remove(selectedMessage);
            messageBindingSource.ResetBindings(false); // Refresh DataGridView
            availableMessages = new List<Message>(allMessages); // Reset available messages when deleting
            // After deletion, clear text boxes for new message entry and disable update/delete buttons
            txtMessageText.Clear();
            nudMessageWeight.Value = 1000;
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            // No automatic save here, rely on explicit Save/Load
            MessageBox.Show("Message deleted successfully.", "Delete Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // Event handler for DataGridView selection changes
    private void DgvMessages_SelectionChanged(object sender, EventArgs e)
    {
        // Enable/disable update/delete buttons based on whether a row is selected
        bool rowSelected = dgvMessages.SelectedRows.Count > 0;
        btnUpdate.Enabled = rowSelected;
        btnDelete.Enabled = rowSelected;

        // Clear the new message input fields when a message is selected in the grid
        // to clearly separate "add" functionality from "edit" functionality.
        txtMessageText.Clear();
        nudMessageWeight.Value = 1000;
    }

    // Event handler for Start/Stop button
    private async void BtnStartStop_Click(object sender, EventArgs e)
    {
        if (!isSending)
        {
            // Get maxLoops value from the NumericUpDown
            int maxLoops = (int)nudLoopCount.Value;
            if (maxLoops <= 0)
            {
                MessageBox.Show("Loop count must be at least 1.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Start sending messages
            btnStartStop.Text = "Stop Sending";
            lblStatus.Text = "Status: Sending messages...";
            isSending = true;
            cancellationTokenSource = new CancellationTokenSource();

            // Disable edit controls while sending is active
            btnAdd.Enabled = false;
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            dgvMessages.Enabled = false;
            btnSaveMessages.Enabled = false; // Disable save/load during sending
            btnLoadMessages.Enabled = false;
            nudLoopCount.Enabled = false; // Disable loop count during sending

            // Run the sending logic in a separate task to keep the UI responsive
            sendingTask = Task.Run(async () => await SendMessagesLoop(cancellationTokenSource.Token, maxLoops));
        }
        else
        {
            // Stop sending messages
            btnStartStop.Text = "Start Sending";
            lblStatus.Text = "Status: Stopping...";
            isSending = false;
            cancellationTokenSource?.Cancel(); // Request cancellation

            // Wait for the task to complete its current iteration or finish
            if (sendingTask != null)
            {
                await sendingTask; // Wait for the task to gracefully finish
            }
            // UI re-enabling is now handled in the finally block of SendMessagesLoop
        }
    }

    // The core message sending loop, now asynchronous and takes maxLoops as argument
    private async Task SendMessagesLoop(CancellationToken cancellationToken, int maxLoops)
    {
        int currentLoop = 0; // Track current loop count
        try // Outer try-catch for OperationCanceledException
        {
            while (!cancellationToken.IsCancellationRequested && currentLoop < maxLoops)
            {
                // If all unique messages have been sent in the current cycle, reset the pool.
                if (availableMessages.Count == 0)
                {
                    // UI update from background thread: use Invoke
                    this.Invoke((MethodInvoker)delegate {
                        lblStatus.Text = $"Status: All unique messages sent. Resetting pool. (Loop {currentLoop + 1}/{maxLoops})";
                    });
                    await Task.Delay(500, cancellationToken); // This can throw TaskCanceledException
                }

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
                    ForceFocusWindow(discordWindow);
                    await Task.Delay(FocusDelayMs, cancellationToken); // Use Task.Delay for non-blocking wait

                    string messageToSend = "";
                    int totalWeightAvailable = availableMessages.Sum(m => m.Weight);

                    if (totalWeightAvailable == 0)
                    {
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: No messages available to send! (Loop {currentLoop + 1}/{maxLoops})";
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
                        messageToSend = selectedMessage.Text;
                        // Remove from available messages on the main thread via Invoke
                        this.Invoke((MethodInvoker)delegate {
                            availableMessages.Remove(selectedMessage);
                            // No need to refresh DGV here, as it's just the 'available' list.
                            // The 'allMessages' list for the DGV remains the master list.
                        });
                    }
                    else
                    {
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Error selecting message. Retrying... (Loop {currentLoop + 1}/{maxLoops})";
                        });
                        await Task.Delay(InputDelayMs * 2, cancellationToken);
                        currentLoop++; // Increment loop even on error
                        continue;
                    }

                    try // Inner try-catch for InputSimulator or SendKeys related errors
                    {
                        string[] lines = messageToSend.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < lines.Length; i++)
                        {
                            inputSimulator.Keyboard.TextEntry(lines[i]);
                            await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException

                            if (i < lines.Length - 1)
                            {
                                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.RETURN);
                                await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException
                            }
                        }

                        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                        await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException

                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Sent message: {messageToSend.Replace("\n", "\\n").Substring(0, Math.Min(50, messageToSend.Length))}... (Loop {currentLoop + 1}/{maxLoops})"; // Truncate for status bar
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // Propagate cancellation from inner block to outer block
                        throw;
                    }
                    catch (Exception ex) // General exception handling for other errors during sending
                    {
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Error with InputSimulator: {ex.Message.Substring(0, Math.Min(50, ex.Message.Length))}. Falling back to SendKeys. (Loop {currentLoop + 1}/{maxLoops})";
                        });

                        string[] lines = messageToSend.Split(new[] { '\n', '\r' }, StringSplitOptions.None);

                        for (int i = 0; i < lines.Length; i++)
                        {
                            SendKeys.SendWait(EscapeSendKeysCharacters(lines[i]));
                            await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException

                            if (i < lines.Length - 1)
                            {
                                SendKeys.SendWait("+{ENTER}");
                                await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException
                            }
                        }
                        SendKeys.SendWait("{ENTER}");
                        await Task.Delay(InputDelayMs, cancellationToken); // This can throw TaskCanceledException
                        this.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = $"Status: Sent (fallback): {messageToSend.Replace("\n", "\\n").Substring(0, Math.Min(50, messageToSend.Length))}... (Loop {currentLoop + 1}/{maxLoops})";
                        });
                    }
                }
                else
                {
                    this.Invoke((MethodInvoker)delegate {
                        lblStatus.Text = "Status: Discord window not found! Please ensure Discord is running.";
                    });
                    // If Discord isn't found, wait longer before re-checking.
                    await Task.Delay(5000, cancellationToken); // This can throw TaskCanceledException
                }
                currentLoop++; // Increment loop count after each attempt (successful or not)

                int delay = random.Next(1000, 7001);
                await Task.Delay(delay, cancellationToken); // This is the line (approx. 632) that caused the original error
            }
        }
        catch (OperationCanceledException) // Catch specific cancellation exception
        {
            // This block will execute when cancellation is requested and an awaitable operation is cancelled.
            // The task will now exit gracefully without throwing an unhandled exception.
            this.Invoke((MethodInvoker)delegate {
                lblStatus.Text = "Status: Stopped by user cancellation.";
            });
        }
        catch (Exception ex) // Catch any other unexpected exceptions that might occur during the loop
        {
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
                dgvMessages.Enabled = true;
                btnSaveMessages.Enabled = true; // Re-enable save/load
                btnLoadMessages.Enabled = true;
                nudLoopCount.Enabled = true; // Re-enable loop count
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
                    lblStatus.Text = $"Status: Messages exported to {Path.GetFileName(saveFileDialog.FileName)}.";
                }
                catch (Exception ex)
                {
                    lblStatus.Text = $"Error exporting messages: {ex.Message}";
                    MessageBox.Show($"Error exporting messages: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                lblStatus.Text = "Status: Export cancelled.";
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

                    if (loadedMessages != null)
                    {
                        allMessages.Clear();
                        allMessages.AddRange(loadedMessages);
                        messageBindingSource.ResetBindings(false); // Refresh DataGridView
                        availableMessages = new List<Message>(allMessages); // Reset available messages after loading
                        lblStatus.Text = $"Status: Messages imported from {Path.GetFileName(openFileDialog.FileName)}.";
                    }
                    else
                    {
                        lblStatus.Text = "Status: Imported file was empty or invalid.";
                        MessageBox.Show("The selected file contains no messages or is in an invalid format.", "Import Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    lblStatus.Text = $"Error importing messages: {ex.Message}";
                    MessageBox.Show($"Error importing messages: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                lblStatus.Text = "Status: Import cancelled.";
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
    private Button btnSave;
    private Button btnCancel;

    public Message EditedMessage { get; private set; }

    // Dark Mode Colors - must be static or passed in
    private static Color _darkBackColor;
    private static Color _darkForeColor;
    private static Color _darkControlBackColor;
    private static Color _darkButtonActiveColor;
    private static Color _dataGridViewHeaderBackColor;


    public EditMessageForm(Message messageToEdit)
    {
        EditedMessage = messageToEdit; // Hold reference to the original message object
        InitializeComponent();
        LoadMessageData();
        ApplyTheme(); // Apply theme after initializing components
    }

    // Method to set theme colors from MainForm
    public void SetThemeColors(Color backColor, Color foreColor, Color controlBackColor, Color buttonActiveColor, Color dataGridViewHeaderBackColor)
    {
        _darkBackColor = backColor;
        _darkForeColor = foreColor;
        _darkControlBackColor = controlBackColor;
        _darkButtonActiveColor = buttonActiveColor;
        _dataGridViewHeaderBackColor = dataGridViewHeaderBackColor;
    }

    private void InitializeComponent()
    {
        this.Text = "Edit Message";
        this.Size = new System.Drawing.Size(500, 300);
        this.MinimumSize = new System.Drawing.Size(350, 250); // Minimum size for the edit form
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.Sizable; // Make the form resizable

        // Message Textbox
        Label lblText = new Label() { Text = "Message Text:", Location = new System.Drawing.Point(10, 10), AutoSize = true };
        txtMessageText = new TextBox()
        {
            Location = new System.Drawing.Point(10, 30),
            Size = new System.Drawing.Size(460, 150),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        // Message Weight NumericUpDown
        Label lblWeight = new Label() { Text = "Weight:", Location = new System.Drawing.Point(10, 190), AutoSize = true };
        nudMessageWeight = new NumericUpDown()
        {
            Location = new System.Drawing.Point(70, 187),
            Size = new System.Drawing.Size(100, 20),
            Minimum = 1,
            Maximum = int.MaxValue, // Max value set to int.MaxValue
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left // Anchor to bottom as form resizes
        };

        // Save Button
        btnSave = new Button()
        {
            Text = "Save",
            Location = new System.Drawing.Point(300, 220),
            Width = 80,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnSave.Click += BtnSave_Click;

        // Cancel Button
        btnCancel = new Button()
        {
            Text = "Cancel",
            Location = new System.Drawing.Point(390, 220),
            Width = 80,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnCancel.Click += BtnCancel_Click;

        this.Controls.Add(lblText);
        this.Controls.Add(txtMessageText);
        this.Controls.Add(lblWeight);
        this.Controls.Add(nudMessageWeight);
        this.Controls.Add(btnSave);
        this.Controls.Add(btnCancel);
    }

    private void ApplyTheme()
    {
        this.BackColor = _darkBackColor;
        this.ForeColor = _darkForeColor;

        ApplyThemeToControls(this.Controls);
    }

    private void ApplyThemeToControls(Control.ControlCollection controls)
    {
        foreach (Control control in controls)
        {
            control.BackColor = _darkControlBackColor;
            control.ForeColor = _darkForeColor;

            if (control is Button button)
            {
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.BackColor = _darkControlBackColor;
                button.ForeColor = _darkForeColor;
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = _darkButtonActiveColor;
                textBox.ForeColor = _darkForeColor;
            }
            else if (control is NumericUpDown numericUpDown)
            {
                numericUpDown.BackColor = _darkButtonActiveColor;
                numericUpDown.ForeColor = _darkForeColor;
            }
            else if (control is Label label)
            {
                label.ForeColor = _darkForeColor;
                label.BackColor = Color.Transparent;
            }

            if (control.HasChildren)
            {
                ApplyThemeToControls(control.Controls);
            }
        }
    }

    private void LoadMessageData()
    {
        if (EditedMessage != null)
        {
            txtMessageText.Text = EditedMessage.Text;
            nudMessageWeight.Value = EditedMessage.Weight;
        }
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        string newText = txtMessageText.Text.Trim();
        int newWeight = (int)nudMessageWeight.Value;

        if (string.IsNullOrEmpty(newText))
        {
            MessageBox.Show("Message text cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Update the properties of the original Message object directly
        EditedMessage.Text = newText;
        EditedMessage.Weight = newWeight;

        this.DialogResult = DialogResult.OK; // Indicate success
        this.Close(); // Close the form
    }

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel; // Indicate cancellation
        this.Close(); // Close the form
    }
}
