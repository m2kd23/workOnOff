using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace WorkResume
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 12. 싱글 인스턴스 실행 보장 (중복 실행 방지)
            bool createdNew;
            using (Mutex mutex = new Mutex(true, "WorkResumeSingleInstanceMutex", out createdNew))
            {
                if (!createdNew)
                {
                    Logger.Write("프로그램이 이미 실행 중입니다. 기존 창을 활성화합니다.", "WARN");
                    Process current = Process.GetCurrentProcess();
                    foreach (Process process in Process.GetProcessesByName(current.ProcessName))
                    {
                        if (process.Id != current.Id)
                        {
                            IntPtr handle = process.MainWindowHandle;
                            if (handle != IntPtr.Zero)
                            {
                                ShowWindow(handle, SW_RESTORE);
                                SetForegroundWindow(handle);
                            }
                            break;
                        }
                    }
                    return;
                }

                Logger.Write("WorkResume 시작");
                Application.Run(new MainForm());
                Logger.Write("WorkResume 종료");
            }
        }
    }

    // 메인 화면 UI
    public class MainForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;
        private const int KEY_1 = 0x31;
        private const int HOTKEY_ID = 9000;
        private const int WM_HOTKEY = 0x0312;

        private Label lblTitle;
        private Label lblOpenedWork;
        private ListView lvOpenedItems;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private Button btnAddFile;
        private Button btnAddFolder;
        private Button btnRemoveSelected;
        
        private Label lblMemo;
        private TextBox txtMemo;
        private Button btnCaptureNow; // "현재 상태로 캡처" (새로 수집)
        private Button btnSaveNextDay; // "다음날 하기" (정확한 UI 버튼명 요구사항 반영)
        private Button btnUpdateSession; // "현재 상태로 업데이트"

        private Label lblSavedSessions;
        private ListView lvSavedSessions;
        private Button btnResume;
        private Button btnViewDetails;
        private Button btnDeleteSession;

        private List<DetectedItem> currentDetectedItems;
        private List<WorkSession> savedSessions;
        private string selectedSessionId = null;

        public MainForm()
        {
            InitializeComponent();
            LoadData();
            RefreshOpenWindows();
        }

        private void InitializeComponent()
        {
            this.Text = "WorkResume - 경량 업무 이어하기";
            this.Size = new Size(820, 720);
            this.MinimumSize = new Size(800, 680);
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 전체 레이아웃 구성을 위한 TableLayoutPanel 사용
            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 5;
            mainLayout.ColumnCount = 1;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));  // 타이틀
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));   // 현재 열린 업무 목록
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110F)); // 업무 메모 & 저장 버튼
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 35F));   // 최근 저장 업무 목록
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10F));  // 여백
            mainLayout.Padding = new Padding(15);
            this.Controls.Add(mainLayout);

            // 1. 타이틀 라벨
            lblTitle = new Label();
            lblTitle.Text = "오늘 하던 업무, 내일 바로 이어하세요";
            lblTitle.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
            lblTitle.Dock = DockStyle.Fill;
            lblTitle.TextAlign = ContentAlignment.MiddleLeft;
            lblTitle.ForeColor = Color.FromArgb(40, 40, 40);
            mainLayout.Controls.Add(lblTitle, 0, 0);

            // 2. 현재 열린 업무 수집 영역 (패널화)
            Panel pnlOpened = new Panel();
            pnlOpened.Dock = DockStyle.Fill;
            mainLayout.Controls.Add(pnlOpened, 0, 1);

            lblOpenedWork = new Label();
            lblOpenedWork.Text = "현재 열려 있는 업무 목록 (체크박스를 선택하세요)";
            lblOpenedWork.Location = new Point(0, 0);
            lblOpenedWork.AutoSize = true;
            lblOpenedWork.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            pnlOpened.Controls.Add(lblOpenedWork);

            lvOpenedItems = new ListView();
            lvOpenedItems.Location = new Point(0, 20);
            lvOpenedItems.Size = new Size(580, 180);
            lvOpenedItems.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lvOpenedItems.View = View.Details;
            lvOpenedItems.CheckBoxes = true;
            lvOpenedItems.FullRowSelect = true;
            lvOpenedItems.GridLines = true;
            lvOpenedItems.Columns.Add("이름", 150);
            lvOpenedItems.Columns.Add("종류", 60);
            lvOpenedItems.Columns.Add("경로", 280);
            lvOpenedItems.Columns.Add("상태", 70);
            pnlOpened.Controls.Add(lvOpenedItems);

            // 현재 열린 목록 조작 버튼들
            FlowLayoutPanel pnlOpenActions = new FlowLayoutPanel();
            pnlOpenActions.Location = new Point(590, 20);
            pnlOpenActions.Size = new Size(180, 180);
            pnlOpenActions.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            pnlOpenActions.FlowDirection = FlowDirection.TopDown;
            pnlOpened.Controls.Add(pnlOpenActions);

            btnSelectAll = CreateButton("전체 선택", new Size(160, 28));
            btnSelectAll.Click += (s, e) => CheckAllOpenedItems(true);
            pnlOpenActions.Controls.Add(btnSelectAll);

            btnDeselectAll = CreateButton("전체 해제", new Size(160, 28));
            btnDeselectAll.Click += (s, e) => CheckAllOpenedItems(false);
            pnlOpenActions.Controls.Add(btnDeselectAll);

            btnAddFile = CreateButton("파일 추가", new Size(160, 28));
            btnAddFile.Click += btnAddFile_Click;
            pnlOpenActions.Controls.Add(btnAddFile);

            btnAddFolder = CreateButton("폴더 추가", new Size(160, 28));
            btnAddFolder.Click += btnAddFolder_Click;
            pnlOpenActions.Controls.Add(btnAddFolder);

            btnRemoveSelected = CreateButton("목록에서 삭제", new Size(160, 28));
            btnRemoveSelected.Click += btnRemoveSelected_Click;
            pnlOpenActions.Controls.Add(btnRemoveSelected);

            btnCaptureNow = CreateButton("현재 상태 다시 캡처", new Size(160, 28));
            btnCaptureNow.Click += (s, e) => RefreshOpenWindows();
            btnCaptureNow.BackColor = Color.FromArgb(240, 240, 240);
            pnlOpenActions.Controls.Add(btnCaptureNow);

            // 3. 메모 및 저장 영역
            Panel pnlSave = new Panel();
            pnlSave.Dock = DockStyle.Fill;
            mainLayout.Controls.Add(pnlSave, 0, 2);

            lblMemo = new Label();
            lblMemo.Text = "업무 메모 (이어할 부분에 대한 간단한 멘트 기록)";
            lblMemo.Location = new Point(0, 5);
            lblMemo.AutoSize = true;
            pnlSave.Controls.Add(lblMemo);

            txtMemo = new TextBox();
            txtMemo.Location = new Point(0, 25);
            txtMemo.Size = new Size(450, 75);
            txtMemo.Multiline = true;
            txtMemo.ScrollBars = ScrollBars.Vertical;
            txtMemo.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pnlSave.Controls.Add(txtMemo);

            btnSaveNextDay = new Button();
            btnSaveNextDay.Text = "다음날 하기"; // (정확히 요구사항 반영)
            btnSaveNextDay.Location = new Point(460, 25);
            btnSaveNextDay.Size = new Size(150, 75);
            btnSaveNextDay.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnSaveNextDay.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
            btnSaveNextDay.BackColor = Color.FromArgb(41, 128, 185);
            btnSaveNextDay.ForeColor = Color.White;
            btnSaveNextDay.FlatStyle = FlatStyle.Flat;
            btnSaveNextDay.FlatAppearance.BorderSize = 0;
            btnSaveNextDay.Click += btnSaveNextDay_Click;
            pnlSave.Controls.Add(btnSaveNextDay);

            btnUpdateSession = new Button();
            btnUpdateSession.Text = "현재 상태로\n업데이트";
            btnUpdateSession.Location = new Point(620, 25);
            btnUpdateSession.Size = new Size(150, 75);
            btnUpdateSession.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnUpdateSession.Font = new Font("맑은 고딕", 11F, FontStyle.Bold);
            btnUpdateSession.BackColor = Color.FromArgb(39, 174, 96);
            btnUpdateSession.ForeColor = Color.White;
            btnUpdateSession.FlatStyle = FlatStyle.Flat;
            btnUpdateSession.FlatAppearance.BorderSize = 0;
            btnUpdateSession.Enabled = false;
            btnUpdateSession.Click += btnUpdateSession_Click;
            pnlSave.Controls.Add(btnUpdateSession);

            // 4. 최근 저장 업무 목록 영역
            Panel pnlSaved = new Panel();
            pnlSaved.Dock = DockStyle.Fill;
            mainLayout.Controls.Add(pnlSaved, 0, 3);

            lblSavedSessions = new Label();
            lblSavedSessions.Text = "최근 저장된 업무 목록";
            lblSavedSessions.Location = new Point(0, 5);
            lblSavedSessions.AutoSize = true;
            lblSavedSessions.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            pnlSaved.Controls.Add(lblSavedSessions);

            lvSavedSessions = new ListView();
            lvSavedSessions.Location = new Point(0, 25);
            lvSavedSessions.Size = new Size(580, 125);
            lvSavedSessions.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lvSavedSessions.View = View.Details;
            lvSavedSessions.FullRowSelect = true;
            lvSavedSessions.GridLines = true;
            lvSavedSessions.MultiSelect = false;
            lvSavedSessions.Columns.Add("저장 일시", 140);
            lvSavedSessions.Columns.Add("업무 메모", 320);
            lvSavedSessions.Columns.Add("파일 수", 60);
            lvSavedSessions.SelectedIndexChanged += lvSavedSessions_SelectedIndexChanged;
            lvSavedSessions.DoubleClick += (s, e) => ResumeSelectedSession();
            pnlSaved.Controls.Add(lvSavedSessions);

            FlowLayoutPanel pnlSavedActions = new FlowLayoutPanel();
            pnlSavedActions.Location = new Point(590, 25);
            pnlSavedActions.Size = new Size(180, 125);
            pnlSavedActions.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            pnlSavedActions.FlowDirection = FlowDirection.TopDown;
            pnlSaved.Controls.Add(pnlSavedActions);

            btnResume = CreateButton("업무 이어하기", new Size(160, 32));
            btnResume.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            btnResume.BackColor = Color.FromArgb(46, 204, 113);
            btnResume.ForeColor = Color.White;
            btnResume.FlatStyle = FlatStyle.Flat;
            btnResume.FlatAppearance.BorderSize = 0;
            btnResume.Click += (s, e) => ResumeSelectedSession();
            pnlSavedActions.Controls.Add(btnResume);

            btnViewDetails = CreateButton("내용 보기", new Size(160, 30));
            btnViewDetails.Click += btnViewDetails_Click;
            pnlSavedActions.Controls.Add(btnViewDetails);

            btnDeleteSession = CreateButton("삭제", new Size(160, 30));
            btnDeleteSession.BackColor = Color.FromArgb(231, 76, 60);
            btnDeleteSession.ForeColor = Color.White;
            btnDeleteSession.FlatStyle = FlatStyle.Flat;
            btnDeleteSession.FlatAppearance.BorderSize = 0;
            btnDeleteSession.Click += btnDeleteSession_Click;
            pnlSavedActions.Controls.Add(btnDeleteSession);
        }

        private Button CreateButton(string text, Size size)
        {
            Button btn = new Button();
            btn.Text = text;
            btn.Size = size;
            btn.Margin = new Padding(0, 0, 0, 6);
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = Color.FromArgb(245, 245, 245);
            btn.FlatAppearance.BorderColor = Color.LightGray;
            return btn;
        }

        private void LoadData()
        {
            savedSessions = SessionManager.LoadSessions();
            RefreshSessionsList();
        }

        private void RefreshSessionsList()
        {
            lvSavedSessions.Items.Clear();
            foreach (var session in savedSessions)
            {
                int fileCount = session.Items.FindAll(i => i.IsSelected).Count;
                ListViewItem item = new ListViewItem(session.SavedAt);
                item.SubItems.Add(session.Description);
                item.SubItems.Add(fileCount.ToString());
                item.Tag = session.Id;
                lvSavedSessions.Items.Add(item);
            }

            // 선택 초기화
            selectedSessionId = null;
            btnUpdateSession.Enabled = false;
        }

        private void RefreshOpenWindows()
        {
            currentDetectedItems = FileDetector.DetectAllOpenedItems();
            PopulateOpenedList();
        }

        private void PopulateOpenedList()
        {
            lvOpenedItems.Items.Clear();
            foreach (var item in currentDetectedItems)
            {
                ListViewItem lvi = new ListViewItem(item.FileName);
                lvi.SubItems.Add(item.Type == "Folder" ? "폴더" : item.Extension.ToUpper().Replace(".", ""));
                lvi.SubItems.Add(item.FilePath);
                lvi.SubItems.Add(item.Exists ? "정상" : "⚠ 누락됨");
                lvi.Checked = item.IsSelected;
                lvi.Tag = item;

                if (!item.Exists)
                {
                    lvi.ForeColor = Color.Red;
                }

                lvOpenedItems.Items.Add(lvi);
            }
        }

        private void CheckAllOpenedItems(bool check)
        {
            foreach (ListViewItem item in lvOpenedItems.Items)
            {
                item.Checked = check;
            }
        }

        private void btnAddFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.Title = "업무 파일 추가";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in ofd.FileNames)
                    {
                        // 중복 검사
                        if (currentDetectedItems.Exists(i => i.FilePath.Equals(file, StringComparison.OrdinalIgnoreCase)))
                            continue;

                        DetectedItem item = new DetectedItem
                        {
                            Type = "File",
                            FilePath = file,
                            FileName = Path.GetFileName(file),
                            Extension = Path.GetExtension(file).ToLower(),
                            PreferredApplication = "",
                            Exists = true,
                            IsSelected = true
                        };
                        currentDetectedItems.Add(item);
                    }
                    PopulateOpenedList();
                    Logger.Write("사용자 수동 파일 추가 완료");
                }
            }
        }

        private void btnAddFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "업무 폴더 추가";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    string folder = fbd.SelectedPath;
                    // 중복 검사
                    if (currentDetectedItems.Exists(i => i.FilePath.Equals(folder, StringComparison.OrdinalIgnoreCase)))
                        return;

                    DetectedItem item = new DetectedItem
                    {
                        Type = "Folder",
                        FilePath = folder,
                        FileName = Path.GetFileName(folder),
                        Extension = "",
                        PreferredApplication = "explorer.exe",
                        Exists = true,
                        IsSelected = true
                    };
                    currentDetectedItems.Add(item);
                    PopulateOpenedList();
                    Logger.Write("사용자 수동 폴더 추가 완료: " + folder);
                }
            }
        }

        private void btnRemoveSelected_Click(object sender, EventArgs e)
        {
            if (lvOpenedItems.SelectedItems.Count == 0)
            {
                MessageBox.Show("목록에서 삭제할 항목을 선택해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (ListViewItem item in lvOpenedItems.SelectedItems)
            {
                DetectedItem det = (DetectedItem)item.Tag;
                currentDetectedItems.Remove(det);
            }
            PopulateOpenedList();
        }

        private void btnSaveNextDay_Click(object sender, EventArgs e)
        {
            SyncCheckedState();
            var selectedItems = currentDetectedItems.FindAll(i => i.IsSelected);
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("저장할 업무 항목이 선택되지 않았습니다.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string memo = txtMemo.Text.Trim();
            string newId = SessionManager.CreateSession(null, memo, currentDetectedItems);
            
            if (newId != null)
            {
                MessageBox.Show(string.Format("현재 업무 {0}개를 저장했습니다.", selectedItems.Count), 
                    "업무 저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                DialogResult closeResult = MessageBox.Show(
                    "저장된 업무에서 사용 중인 프로그램들을 모두 종료하시겠습니까?\n" +
                    "(저장되지 않은 작업이 있는 경우 종료 중 저장할 수 있는 창이 뜹니다.)",
                    "프로그램 종료 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (closeResult == DialogResult.Yes)
                {
                    CloseSavedApplications(selectedItems);
                }

                txtMemo.Clear();
                LoadData();
                RefreshOpenWindows();
            }
        }

        private void CloseSavedApplications(List<DetectedItem> savedItems)
        {
            Logger.Write("저장 업무 프로그램 종료 처리 시작");
            
            bool closedExcel = false;
            bool closedWord = false;
            bool closedPpt = false;
            bool closedHwp = false;

            foreach (var item in savedItems)
            {
                if (item.Type != "File") continue;
                string ext = item.Extension.ToLower();

                if ((ext == ".xlsx" || ext == ".xls" || ext == ".xlsm") && !closedExcel)
                {
                    try
                    {
                        dynamic excel = Marshal.GetActiveObject("Excel.Application");
                        if (excel != null)
                        {
                            excel.Quit();
                            closedExcel = true;
                            Logger.Write("Excel Application Quit 요청 송신");
                        }
                    }
                    catch {}
                }
                else if ((ext == ".docx" || ext == ".doc") && !closedWord)
                {
                    try
                    {
                        dynamic word = Marshal.GetActiveObject("Word.Application");
                        if (word != null)
                        {
                            word.Quit();
                            closedWord = true;
                            Logger.Write("Word Application Quit 요청 송신");
                        }
                    }
                    catch {}
                }
                else if ((ext == ".pptx" || ext == ".ppt") && !closedPpt)
                {
                    try
                    {
                        dynamic ppt = Marshal.GetActiveObject("PowerPoint.Application");
                        if (ppt != null)
                        {
                            ppt.Quit();
                            closedPpt = true;
                            Logger.Write("PowerPoint Application Quit 요청 송신");
                        }
                    }
                    catch {}
                }
                else if ((ext == ".hwp" || ext == ".hwpx") && !closedHwp)
                {
                    try
                    {
                        dynamic hwp = Marshal.GetActiveObject("HWPFrame.HwpObject");
                        if (hwp != null)
                        {
                            hwp.Quit();
                            closedHwp = true;
                            Logger.Write("한글(HWP) Application Quit 요청 송신");
                        }
                    }
                    catch {}
                }
            }

            foreach (var item in savedItems)
            {
                if (item.Type != "File") continue;
                string fileName = item.FileName;
                string ext = item.Extension.ToLower();

                if (ext == ".txt" || ext == ".csv")
                {
                    CloseProcessesByNameAndTitle("notepad", fileName);
                }
                else if (ext == ".pdf")
                {
                    CloseProcessesByNameAndTitle("AcroRd32", fileName);
                    CloseProcessesByNameAndTitle("Acrobat", fileName);
                    CloseProcessesByNameAndTitle("msedge", fileName);
                    CloseProcessesByNameAndTitle("chrome", fileName);
                }
            }
            
            CloseSavedExplorerWindows(savedItems);
        }

        private void CloseProcessesByNameAndTitle(string processName, string fileName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(processName);
                foreach (Process proc in processes)
                {
                    try
                    {
                        string title = proc.MainWindowTitle;
                        if (!string.IsNullOrEmpty(title) && title.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Logger.Write(string.Format("프로세스 {0} (창 제목: {1}) CloseMainWindow 요청 송신", processName, title));
                            proc.CloseMainWindow();
                        }
                    }
                    catch {}
                }
            }
            catch {}
        }

        private void CloseSavedExplorerWindows(List<DetectedItem> savedItems)
        {
            try
            {
                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null) return;
                dynamic shell = Activator.CreateInstance(shellType);
                dynamic windows = shell.Windows();
                if (windows == null) return;

                int count = windows.Count;
                for (int i = count - 1; i >= 0; i--)
                {
                    try
                    {
                        dynamic window = windows.Item(i);
                        if (window == null) continue;

                        string fullName = window.FullName;
                        if (fullName != null && fullName.ToLower().EndsWith("explorer.exe"))
                        {
                            string locationUrl = window.LocationURL;
                            if (!string.IsNullOrEmpty(locationUrl))
                            {
                                Uri uri = new Uri(locationUrl);
                                if (uri.IsFile)
                                {
                                    string localPath = uri.LocalPath;
                                    bool isSavedFolder = savedItems.Exists(x => x.Type == "Folder" && x.FilePath.Equals(localPath, StringComparison.OrdinalIgnoreCase));
                                    if (isSavedFolder)
                                    {
                                        Logger.Write("탐색기 창 닫기 요청 송신: " + localPath);
                                        window.Quit();
                                    }
                                }
                            }
                        }
                    }
                    catch {}
                }
            }
            catch {}
        }

        private void btnUpdateSession_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedSessionId)) return;

            SyncCheckedState();
            var selectedItems = currentDetectedItems.FindAll(i => i.IsSelected);
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("저장할 업무 항목이 선택되지 않았습니다.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string memo = txtMemo.Text.Trim();
            bool success = SessionManager.UpdateSession(selectedSessionId, null, memo, currentDetectedItems);

            if (success)
            {
                MessageBox.Show("선택한 세션의 작업 목록과 메모를 현재 상태로 업데이트했습니다.", 
                    "업데이트 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMemo.Clear();
                LoadData();
                RefreshOpenWindows();
            }
        }

        private void SyncCheckedState()
        {
            foreach (ListViewItem lvi in lvOpenedItems.Items)
            {
                DetectedItem det = (DetectedItem)lvi.Tag;
                det.IsSelected = lvi.Checked;
            }
        }

        private void lvSavedSessions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvSavedSessions.SelectedItems.Count > 0)
            {
                selectedSessionId = (string)lvSavedSessions.SelectedItems[0].Tag;
                btnUpdateSession.Enabled = true;
                
                // 메모 가져와서 메모 박스에 로드
                WorkSession s = savedSessions.Find(x => x.Id == selectedSessionId);
                if (s != null)
                {
                    txtMemo.Text = s.Description;
                }
            }
            else
            {
                selectedSessionId = null;
                btnUpdateSession.Enabled = false;
            }
        }

        private void ResumeSelectedSession()
        {
            if (lvSavedSessions.SelectedItems.Count == 0)
            {
                MessageBox.Show("이어하기를 할 업무 세션을 선택해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string sessionId = (string)lvSavedSessions.SelectedItems[0].Tag;
            WorkSession session = savedSessions.Find(s => s.Id == sessionId);
            if (session == null) return;

            Logger.Write("업무 이어하기 시작: " + session.Title);
            
            List<string> missingItems = new List<string>();
            int successCount = 0;

            foreach (var item in session.Items)
            {
                if (!item.IsSelected) continue;

                // 4-6. 파일 존재 여부 확인
                if (item.Type == "File")
                {
                    if (!File.Exists(item.FilePath))
                    {
                        missingItems.Add("[파일 없음] " + item.FilePath);
                        continue;
                    }
                }
                else if (item.Type == "Folder")
                {
                    if (!Directory.Exists(item.FilePath))
                    {
                        missingItems.Add("[폴더 없음] " + item.FilePath);
                        continue;
                    }
                }

                // 실행 시도 (ShellExecute 기반)
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = item.FilePath,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    successCount++;
                }
                catch (Exception ex)
                {
                    missingItems.Add(string.Format("[실행 실패] {0} (이유: {1})", item.FilePath, ex.Message));
                    Logger.Write(string.Format("파일 재실행 실패: {0}, 에러: {1}", item.FilePath, ex.Message), "ERROR");
                }
            }

            // 누락/오류 목록 표시 (비차단 예외 처리 완료)
            if (missingItems.Count > 0)
            {
                string msg = "일부 파일을 열지 못했습니다.\n\n" + string.Join("\n", missingItems);
                MessageBox.Show(msg, "업무 복원 경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.Write("일부 파일 복원 누락 발생", "WARN");
            }
            else
            {
                MessageBox.Show(string.Format("업무 복원이 완료되었습니다. ({0}개 항목 복구 완료)", successCount), 
                    "업무 이어하기 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnDeleteSession_Click(object sender, EventArgs e)
        {
            if (lvSavedSessions.SelectedItems.Count == 0)
            {
                MessageBox.Show("삭제할 업무 세션을 선택해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string sessionId = (string)lvSavedSessions.SelectedItems[0].Tag;
            WorkSession session = savedSessions.Find(s => s.Id == sessionId);
            if (session == null) return;

            var result = MessageBox.Show(string.Format("'{0}' 세션을 삭제하시겠습니까?", session.Title), 
                "세션 삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                if (SessionManager.DeleteSession(sessionId))
                {
                    LoadData();
                }
            }
        }

        private void btnViewDetails_Click(object sender, EventArgs e)
        {
            if (lvSavedSessions.SelectedItems.Count == 0)
            {
                MessageBox.Show("상세 보기를 할 업무 세션을 선택해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string sessionId = (string)lvSavedSessions.SelectedItems[0].Tag;
            WorkSession session = savedSessions.Find(s => s.Id == sessionId);
            if (session != null)
            {
                using (DetailsForm detailsForm = new DetailsForm(session, () => LoadData()))
                {
                    detailsForm.ShowDialog(this);
                }
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            bool success = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, KEY_1);
            if (success)
            {
                Logger.Write("전역 단축키 Ctrl+Shift+1 등록 성공");
            }
            else
            {
                Logger.Write("전역 단축키 Ctrl+Shift+1 등록 실패 (이미 다른 프로그램이 사용 중)", "WARN");
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosing(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                ActivateMainWindow();
            }
            base.WndProc(ref m);
        }

        private void ActivateMainWindow()
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            this.Show();
            this.Activate();
            SetForegroundWindow(this.Handle);
            Logger.Write("전역 단축키 입력으로 창 활성화");
        }
    }

    // 세션 상세 보기 UI
    public class DetailsForm : Form
    {
        private WorkSession session;
        private Action onDataChanged;
        private ListView lvItems;
        private Button btnOpenSingle;
        private Button btnOpenFolder;
        private Button btnDeleteFromSession;
        private Button btnClose;

        public DetailsForm(WorkSession session, Action onDataChanged)
        {
            this.session = session;
            this.onDataChanged = onDataChanged;
            InitializeComponent();
            PopulateList();
        }

        private void InitializeComponent()
        {
            this.Text = string.Format("세션 상세 보기 - {0}", session.Title);
            this.Size = new Size(700, 450);
            this.MinimumSize = new Size(600, 400);
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.StartPosition = FormStartPosition.CenterParent;

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.RowCount = 3;
            layout.ColumnCount = 1;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));  // 정보 라벨
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 파일 리스트
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));  // 액션 버튼 영역
            layout.Padding = new Padding(12);
            this.Controls.Add(layout);

            Label lblInfo = new Label();
            lblInfo.Text = string.Format("저장 일시: {0}  |  메모: {1}", session.SavedAt, session.Description);
            lblInfo.Dock = DockStyle.Fill;
            lblInfo.TextAlign = ContentAlignment.MiddleLeft;
            lblInfo.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            layout.Controls.Add(lblInfo, 0, 0);

            lvItems = new ListView();
            lvItems.Dock = DockStyle.Fill;
            lvItems.View = View.Details;
            lvItems.FullRowSelect = true;
            lvItems.GridLines = true;
            lvItems.MultiSelect = false;
            lvItems.Columns.Add("이름", 150);
            lvItems.Columns.Add("종류", 60);
            lvItems.Columns.Add("경로", 320);
            lvItems.Columns.Add("상태", 70);
            layout.Controls.Add(lvItems, 0, 1);

            Panel pnlActions = new Panel();
            pnlActions.Dock = DockStyle.Fill;
            layout.Controls.Add(pnlActions, 0, 2);

            btnOpenSingle = new Button();
            btnOpenSingle.Text = "열기";
            btnOpenSingle.Location = new Point(0, 8);
            btnOpenSingle.Size = new Size(100, 30);
            btnOpenSingle.FlatStyle = FlatStyle.Flat;
            btnOpenSingle.Click += btnOpenSingle_Click;
            pnlActions.Controls.Add(btnOpenSingle);

            btnOpenFolder = new Button();
            btnOpenFolder.Text = "폴더 열기";
            btnOpenFolder.Location = new Point(106, 8);
            btnOpenFolder.Size = new Size(100, 30);
            btnOpenFolder.FlatStyle = FlatStyle.Flat;
            btnOpenFolder.Click += btnOpenFolder_Click;
            pnlActions.Controls.Add(btnOpenFolder);

            btnDeleteFromSession = new Button();
            btnDeleteFromSession.Text = "목록에서 삭제";
            btnDeleteFromSession.Location = new Point(212, 8);
            btnDeleteFromSession.Size = new Size(120, 30);
            btnDeleteFromSession.FlatStyle = FlatStyle.Flat;
            btnDeleteFromSession.BackColor = Color.FromArgb(231, 76, 60);
            btnDeleteFromSession.ForeColor = Color.White;
            btnDeleteFromSession.FlatAppearance.BorderSize = 0;
            btnDeleteFromSession.Click += btnDeleteFromSession_Click;
            pnlActions.Controls.Add(btnDeleteFromSession);

            btnClose = new Button();
            btnClose.Text = "닫기";
            btnClose.Location = new Point(576, 8);
            btnClose.Size = new Size(100, 30);
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.Click += (s, e) => this.Close();
            pnlActions.Controls.Add(btnClose);
        }

        private void PopulateList()
        {
            lvItems.Items.Clear();
            foreach (var item in session.Items)
            {
                ListViewItem lvi = new ListViewItem(item.FileName);
                lvi.SubItems.Add(item.Type == "Folder" ? "폴더" : item.Extension.ToUpper().Replace(".", ""));
                lvi.SubItems.Add(item.FilePath);
                lvi.SubItems.Add(item.Exists ? "정상" : "⚠ 누락됨");
                lvi.Tag = item;

                if (!item.Exists)
                {
                    lvi.ForeColor = Color.Red;
                }

                lvItems.Items.Add(lvi);
            }
        }

        private void btnOpenSingle_Click(object sender, EventArgs e)
        {
            if (lvItems.SelectedItems.Count == 0) return;
            SessionItem item = (SessionItem)lvItems.SelectedItems[0].Tag;

            if (item.Type == "File" && !File.Exists(item.FilePath))
            {
                MessageBox.Show("파일이 존재하지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (item.Type == "Folder" && !Directory.Exists(item.FilePath))
            {
                MessageBox.Show("폴더가 존재하지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = item.FilePath,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("실행 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            if (lvItems.SelectedItems.Count == 0) return;
            SessionItem item = (SessionItem)lvItems.SelectedItems[0].Tag;

            string folder = item.Type == "Folder" ? item.FilePath : Path.GetDirectoryName(item.FilePath);
            if (!Directory.Exists(folder))
            {
                MessageBox.Show("폴더가 존재하지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (item.Type == "File" && File.Exists(item.FilePath))
                {
                    // 파일을 탐색기에서 선택된 상태로 띄움
                    Process.Start("explorer.exe", string.Format("/select,\"{0}\"", item.FilePath));
                }
                else
                {
                    Process.Start("explorer.exe", string.Format("\"{0}\"", folder));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("폴더 열기 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteFromSession_Click(object sender, EventArgs e)
        {
            if (lvItems.SelectedItems.Count == 0) return;
            SessionItem item = (SessionItem)lvItems.SelectedItems[0].Tag;

            var result = MessageBox.Show(string.Format("'{0}' 항목을 이 세션에서 삭제하시겠습니까?", item.FileName),
                "항목 삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                session.Items.Remove(item);
                
                // sessions.json 업데이트
                List<WorkSession> all = SessionManager.LoadSessions();
                int idx = all.FindIndex(s => s.Id == session.Id);
                if (idx >= 0)
                {
                    all[idx] = session;
                    SessionManager.SaveSessions(all);
                }

                PopulateList();
                if (onDataChanged != null)
                {
                    onDataChanged();
                }
            }
        }
    }
}
