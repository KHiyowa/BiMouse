using Microsoft.Win32;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace BiMouse
{
    public partial class Tasktray : Component
    {
        [DllImport("user32.dll")]
        private static extern int SwapMouseButton(int bSwap);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;
        private const int VK_RBUTTON = 0x02;
        private const int SM_SWAPBUTTON = 23;

        private IconWatchdog watchdog;
        private NotifyIcon notifyIcon;
        private ToolStripMenuItem rightHandItem;
        private ToolStripMenuItem leftHandItem;
        private Icon rightIcon;
        private Icon leftIcon;

        private System.Windows.Forms.Timer _pollingTimer;
        private DateTime? _bothDownStart;
        private bool _triggered;

        public Tasktray()
        {
            this.SetComponents();
            InitializeComponent();

            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;

            Application.ApplicationExit += (s, e) =>
            {
                SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
                if (watchdog != null) watchdog.ReleaseHandle();
            };
        }

        private void SetComponents()
        {
            notifyIcon = new NotifyIcon();

            string rightPath = Path.Combine(Application.StartupPath, "resources/icons/bimouse_right.ico");
            string leftPath = Path.Combine(Application.StartupPath, "resources/icons/bimouse_left.ico");

            try { rightIcon = File.Exists(rightPath) ? new Icon(rightPath) : SystemIcons.Application; } catch { rightIcon = SystemIcons.Application; }
            try { leftIcon = File.Exists(leftPath) ? new Icon(leftPath) : SystemIcons.Application; } catch { leftIcon = SystemIcons.Application; }

            notifyIcon.Visible = true;
            notifyIcon.Text = Application.ProductName;
            notifyIcon.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.NonPublic | BindingFlags.Instance);
                    mi?.Invoke(notifyIcon, null);
                }
            };

            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version?.ToString() ?? "N/A";
            ToolStripMenuItem versionItem = new ToolStripMenuItem($"{Application.ProductName} v{version}");
            versionItem.Enabled = false;

            ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

            rightHandItem = new ToolStripMenuItem();
            rightHandItem.Text = "右手マウス";
            rightHandItem.Click += RightHand_Click;
            rightHandItem.MouseUp += MenuItem_MouseUp;

            leftHandItem = new ToolStripMenuItem();
            leftHandItem.Text = "左手マウス";
            leftHandItem.Click += LeftHand_Click;
            leftHandItem.MouseUp += MenuItem_MouseUp;

            ToolStripMenuItem exit = new ToolStripMenuItem();
            exit.Text = "終了(Ctrl + クリック)";
            exit.Click += Exit_Click;
            exit.MouseUp += MenuItem_MouseUp;

            contextMenuStrip.Items.AddRange(new ToolStripItem[] {
                versionItem,
                new ToolStripSeparator(),
                rightHandItem,
                leftHandItem,
                new ToolStripSeparator(),
                exit 
            });
            notifyIcon.ContextMenuStrip = contextMenuStrip;

            UpdateMenuState();

            // 監視用タイマー(Polling)
            _pollingTimer = new System.Windows.Forms.Timer();
            _pollingTimer.Interval = 50;
            _pollingTimer.Tick += PollingTimer_Tick;
            _pollingTimer.Start();

            watchdog = new IconWatchdog(() =>
            {
                if (notifyIcon != null)
                {
                    notifyIcon.Visible = false;
                    notifyIcon.Visible = true;
                }
            });
        }

        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Mouse)
            {
                UpdateMenuState();
            }
        }

        private void UpdateMenuState()
        {
            bool isSwapped = GetSystemMetrics(SM_SWAPBUTTON) != 0;
            
            rightHandItem.Checked = !isSwapped;
            rightHandItem.Enabled = isSwapped;

            leftHandItem.Checked = isSwapped;
            leftHandItem.Enabled = !isSwapped;

            notifyIcon.Text = isSwapped ? $"{Application.ProductName} (左手)" : $"{Application.ProductName} (右手)";
            notifyIcon.Icon = isSwapped ? leftIcon : rightIcon;
        }

        private void PollingTimer_Tick(object sender, EventArgs e)
        {
            bool left = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
            bool right = (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;

            if (left && right)
            {
                if (_bothDownStart == null)
                {
                    _bothDownStart = DateTime.Now;
                }
                else if (!_triggered && (DateTime.Now - _bothDownStart.Value).TotalMilliseconds >= 1000)
                {
                    _triggered = true;
                    bool isSwapped = GetSystemMetrics(SM_SWAPBUTTON) != 0;
                    bool newState = !isSwapped;
                    SwapMouseButton(newState ? 1 : 0);
                    UpdateMenuState();
                    string stateText = newState ? "左手マウス" : "右手マウス";
                    notifyIcon.ShowBalloonTip(1000, Application.ProductName, $"{stateText}に切り替えました", ToolTipIcon.Info);
                }
            }
            else
            {
                _bothDownStart = null;
                _triggered = false;
            }
        }

        private void RightHand_Click(object sender, EventArgs e)
        {
            SwapMouseButton(0);
            UpdateMenuState();
        }

        private void LeftHand_Click(object sender, EventArgs e)
        {
            SwapMouseButton(1);
            UpdateMenuState();
        }

        private void MenuItem_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                (sender as ToolStripMenuItem)?.PerformClick();
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            if ((Control.ModifierKeys & Keys.Control) != Keys.Control)
            {
                notifyIcon.ShowBalloonTip(1000, Application.ProductName, "終了するには Ctrl キーを押しながらクリックしてください", ToolTipIcon.Info);
                return;
            }
            Application.Exit();
        }

        private class IconWatchdog : NativeWindow
        {
            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            static extern uint RegisterWindowMessage(string lpString);
            private uint wmTaskbarCreated;
            private Action onTaskbarCreated;

            public IconWatchdog(Action callback)
            {
                onTaskbarCreated = callback;
                wmTaskbarCreated = RegisterWindowMessage("TaskbarCreated");
                this.CreateHandle(new CreateParams());
            }

            protected override void WndProc(ref Message m)
            {
                if (wmTaskbarCreated != 0 && m.Msg == wmTaskbarCreated)
                {
                    onTaskbarCreated?.Invoke();
                }
                base.WndProc(ref m);
            }
        }
    }
}