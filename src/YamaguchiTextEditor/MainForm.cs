using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// テキストエディタのメインフォーム
    /// </summary>
    public class MainForm : Form
    {
        // コントロール
        private MenuStrip menuStrip;
        private ToolStrip toolStrip;
        private StatusStrip statusStrip;
        private Panel panelEditor;
        private LineNumberPanel lineNumberPanel;
        private SpecialCharRichTextBox textEditor;  // 特殊文字表示対応
        
        // 表モード用
        private TableModePanel tableModePanel;
        private bool isTableMode;
        private ToolStripButton btnTableMode;
        private ToolStripMenuItem menuTableMode;
        private ToolStripMenuItem menuShowSpecialChars;  // 特殊文字表示メニュー

        // ステータスバーラベル
        private ToolStripStatusLabel lblMode;
        private ToolStripStatusLabel lblPosition;
        private ToolStripStatusLabel lblSelection;
        private ToolStripStatusLabel lblCharCode;  // 文字コード表示
        private ToolStripStatusLabel lblEncoding;
        private ToolStripStatusLabel lblNewLine;

        // メニュー項目
        private ToolStripMenuItem menuFile;
        private ToolStripMenuItem menuEdit;
        private ToolStripMenuItem menuSearch;
        private ToolStripMenuItem menuView;
        private ToolStripMenuItem menuFormat;
        private ToolStripMenuItem menuHelp;

        // ダイアログ
        private SearchForm searchForm;
        private ColumnSearchForm columnSearchForm;

        // 設定
        private EditorSettings settings;

        // ステータスバー更新用タイマー
        private Timer statusUpdateTimer;
        private bool statusUpdateRequested;

        // アプリケーション名
        private const string AppName = "Yamaguchi Text Editor";

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainForm(string initialFilePath = null)
        {
            settings = new EditorSettings();
            isTableMode = false;
            
            InitializeComponents();
            InitializeMenuStrip();
            InitializeToolStrip();
            InitializeStatusStrip();
            InitializeEditorPanel();
            InitializeTableModePanel();
            
            UpdateTitle();
            UpdateStatusBar();

            // 起動時にファイルが指定されていれば開く
            if (!string.IsNullOrEmpty(initialFilePath) && File.Exists(initialFilePath))
            {
                OpenFile(initialFilePath, settings.CurrentEncoding);
            }
        }

        /// <summary>
        /// フォーム初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Text = AppName;
            this.Size = new Size(1024, 768);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("MS UI Gothic", 9.0f);
            this.KeyPreview = true;
            this.KeyDown += MainForm_KeyDown;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);

            // ドラッグ＆ドロップを有効化（フォーム）
            this.AllowDrop = true;
            this.DragEnter += Control_DragEnter;
            this.DragDrop += Control_DragDrop;

            // ステータスバー更新タイマー初期化
            statusUpdateTimer = new Timer();
            statusUpdateTimer.Interval = 100;  // 100ms間隔
            statusUpdateTimer.Tick += StatusUpdateTimer_Tick;
            statusUpdateRequested = false;
        }

        /// <summary>
        /// ステータスバー更新タイマーTick
        /// </summary>
        private void StatusUpdateTimer_Tick(object sender, EventArgs e)
        {
            statusUpdateTimer.Stop();
            if (statusUpdateRequested)
            {
                statusUpdateRequested = false;
                UpdateStatusBarInternal();
            }
        }

        /// <summary>
        /// ステータスバー更新をリクエスト（遅延実行）
        /// </summary>
        private void RequestStatusUpdate()
        {
            statusUpdateRequested = true;
            if (!statusUpdateTimer.Enabled)
            {
                statusUpdateTimer.Start();
            }
        }

        /// <summary>
        /// メニューストリップ初期化
        /// </summary>
        private void InitializeMenuStrip()
        {
            menuStrip = new MenuStrip();
            menuStrip.Dock = DockStyle.Top;

            // ファイルメニュー
            menuFile = new ToolStripMenuItem("ファイル(&F)");
            
            var menuNew = new ToolStripMenuItem("新規作成(&N)", null, Menu_New_Click);
            menuNew.ShortcutKeys = Keys.Control | Keys.N;
            menuFile.DropDownItems.Add(menuNew);

            var menuOpen = new ToolStripMenuItem("開く(&O)...", null, Menu_Open_Click);
            menuOpen.ShortcutKeys = Keys.Control | Keys.O;
            menuFile.DropDownItems.Add(menuOpen);

            var menuSave = new ToolStripMenuItem("上書き保存(&S)", null, Menu_Save_Click);
            menuSave.ShortcutKeys = Keys.Control | Keys.S;
            menuFile.DropDownItems.Add(menuSave);

            var menuSaveAs = new ToolStripMenuItem("名前を付けて保存(&A)...", null, Menu_SaveAs_Click);
            menuSaveAs.ShortcutKeys = Keys.Control | Keys.Shift | Keys.S;
            menuFile.DropDownItems.Add(menuSaveAs);

            menuFile.DropDownItems.Add(new ToolStripSeparator());

            var menuExit = new ToolStripMenuItem("終了(&X)", null, Menu_Exit_Click);
            menuFile.DropDownItems.Add(menuExit);

            menuStrip.Items.Add(menuFile);

            // 編集メニュー
            menuEdit = new ToolStripMenuItem("編集(&E)");

            var menuUndo = new ToolStripMenuItem("元に戻す(&U)", null, Menu_Undo_Click);
            menuUndo.ShortcutKeys = Keys.Control | Keys.Z;
            menuEdit.DropDownItems.Add(menuUndo);

            var menuRedo = new ToolStripMenuItem("やり直し(&R)", null, Menu_Redo_Click);
            menuRedo.ShortcutKeys = Keys.Control | Keys.Y;
            menuEdit.DropDownItems.Add(menuRedo);

            menuEdit.DropDownItems.Add(new ToolStripSeparator());

            var menuCut = new ToolStripMenuItem("切り取り(&T)", null, Menu_Cut_Click);
            menuCut.ShortcutKeys = Keys.Control | Keys.X;
            menuEdit.DropDownItems.Add(menuCut);

            var menuCopy = new ToolStripMenuItem("コピー(&C)", null, Menu_Copy_Click);
            menuCopy.ShortcutKeys = Keys.Control | Keys.C;
            menuEdit.DropDownItems.Add(menuCopy);

            var menuPaste = new ToolStripMenuItem("貼り付け(&P)", null, Menu_Paste_Click);
            menuPaste.ShortcutKeys = Keys.Control | Keys.V;
            menuEdit.DropDownItems.Add(menuPaste);

            menuEdit.DropDownItems.Add(new ToolStripSeparator());

            var menuSelectAll = new ToolStripMenuItem("すべて選択(&A)", null, Menu_SelectAll_Click);
            menuSelectAll.ShortcutKeys = Keys.Control | Keys.A;
            menuEdit.DropDownItems.Add(menuSelectAll);

            menuStrip.Items.Add(menuEdit);

            // 表示メニュー
            menuView = new ToolStripMenuItem("表示(&V)");

            menuTableMode = new ToolStripMenuItem("表モード(&T)", null, Menu_TableMode_Click);
            menuTableMode.ShortcutKeys = Keys.Control | Keys.T;
            menuTableMode.CheckOnClick = true;
            menuView.DropDownItems.Add(menuTableMode);

            menuView.DropDownItems.Add(new ToolStripSeparator());

            menuShowSpecialChars = new ToolStripMenuItem("特殊文字を表示(&S)", null, Menu_ShowSpecialChars_Click);
            menuShowSpecialChars.ShortcutKeys = Keys.Control | Keys.Shift | Keys.H;
            menuShowSpecialChars.CheckOnClick = true;
            menuShowSpecialChars.Checked = true;  // デフォルトでチェックON
            menuView.DropDownItems.Add(menuShowSpecialChars);

            menuStrip.Items.Add(menuView);

            // 検索メニュー
            menuSearch = new ToolStripMenuItem("検索(&S)");

            var menuFind = new ToolStripMenuItem("検索(&F)...", null, Menu_Find_Click);
            menuFind.ShortcutKeys = Keys.Control | Keys.F;
            menuSearch.DropDownItems.Add(menuFind);

            var menuFindNext = new ToolStripMenuItem("次を検索(&N)", null, Menu_FindNext_Click);
            menuFindNext.ShortcutKeys = Keys.F3;
            menuSearch.DropDownItems.Add(menuFindNext);

            var menuFindPrev = new ToolStripMenuItem("前を検索(&P)", null, Menu_FindPrev_Click);
            menuFindPrev.ShortcutKeys = Keys.Shift | Keys.F3;
            menuSearch.DropDownItems.Add(menuFindPrev);

            var menuReplace = new ToolStripMenuItem("置換(&R)...", null, Menu_Replace_Click);
            menuReplace.ShortcutKeys = Keys.Control | Keys.H;
            menuSearch.DropDownItems.Add(menuReplace);

            menuSearch.DropDownItems.Add(new ToolStripSeparator());

            var menuGoToLine = new ToolStripMenuItem("指定行へ移動(&G)...", null, Menu_GoToLine_Click);
            menuGoToLine.ShortcutKeys = Keys.Control | Keys.G;
            menuSearch.DropDownItems.Add(menuGoToLine);

            menuSearch.DropDownItems.Add(new ToolStripSeparator());

            var menuColumnSearch = new ToolStripMenuItem("列の検索・置換(&L)...", null, Menu_ColumnSearch_Click);
            menuColumnSearch.ShortcutKeys = Keys.Control | Keys.L;
            menuSearch.DropDownItems.Add(menuColumnSearch);

            menuStrip.Items.Add(menuSearch);

            // 書式メニュー
            menuFormat = new ToolStripMenuItem("書式(&O)");

            var menuFont = new ToolStripMenuItem("フォント(&F)...", null, Menu_Font_Click);
            menuFormat.DropDownItems.Add(menuFont);

            menuFormat.DropDownItems.Add(new ToolStripSeparator());

            var menuEncoding = new ToolStripMenuItem("文字コード(&E)");
            
            var menuEncodingUtf8 = new ToolStripMenuItem("UTF-8", null, Menu_Encoding_Click);
            menuEncodingUtf8.Tag = new UTF8Encoding(false);
            menuEncoding.DropDownItems.Add(menuEncodingUtf8);

            var menuEncodingUtf8Bom = new ToolStripMenuItem("UTF-8 (BOM付き)", null, Menu_Encoding_Click);
            menuEncodingUtf8Bom.Tag = new UTF8Encoding(true);
            menuEncoding.DropDownItems.Add(menuEncodingUtf8Bom);

            var menuEncodingShiftJis = new ToolStripMenuItem("Shift-JIS", null, Menu_Encoding_Click);
            menuEncodingShiftJis.Tag = Encoding.GetEncoding(932);
            menuEncoding.DropDownItems.Add(menuEncodingShiftJis);

            var menuEncodingEucJp = new ToolStripMenuItem("EUC-JP", null, Menu_Encoding_Click);
            menuEncodingEucJp.Tag = Encoding.GetEncoding(51932);
            menuEncoding.DropDownItems.Add(menuEncodingEucJp);

            var menuEncodingJis = new ToolStripMenuItem("JIS", null, Menu_Encoding_Click);
            menuEncodingJis.Tag = Encoding.GetEncoding(50220);
            menuEncoding.DropDownItems.Add(menuEncodingJis);

            var menuEncodingUtf16Le = new ToolStripMenuItem("UTF-16 LE", null, Menu_Encoding_Click);
            menuEncodingUtf16Le.Tag = Encoding.Unicode;
            menuEncoding.DropDownItems.Add(menuEncodingUtf16Le);

            var menuEncodingUtf16Be = new ToolStripMenuItem("UTF-16 BE", null, Menu_Encoding_Click);
            menuEncodingUtf16Be.Tag = Encoding.BigEndianUnicode;
            menuEncoding.DropDownItems.Add(menuEncodingUtf16Be);

            menuFormat.DropDownItems.Add(menuEncoding);

            menuStrip.Items.Add(menuFormat);

            // ヘルプメニュー
            menuHelp = new ToolStripMenuItem("ヘルプ(&H)");

            var menuAbout = new ToolStripMenuItem("バージョン情報(&A)", null, Menu_About_Click);
            menuHelp.DropDownItems.Add(menuAbout);

            menuStrip.Items.Add(menuHelp);

            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        /// <summary>
        /// ツールストリップ初期化
        /// </summary>
        private void InitializeToolStrip()
        {
            toolStrip = new ToolStrip();
            toolStrip.Dock = DockStyle.Top;

            var btnNew = new ToolStripButton("新規");
            btnNew.ToolTipText = "新規作成 (Ctrl+N)";
            btnNew.Click += Menu_New_Click;
            toolStrip.Items.Add(btnNew);

            var btnOpen = new ToolStripButton("開く");
            btnOpen.ToolTipText = "ファイルを開く (Ctrl+O)";
            btnOpen.Click += Menu_Open_Click;
            toolStrip.Items.Add(btnOpen);

            var btnSave = new ToolStripButton("保存");
            btnSave.ToolTipText = "上書き保存 (Ctrl+S)";
            btnSave.Click += Menu_Save_Click;
            toolStrip.Items.Add(btnSave);

            toolStrip.Items.Add(new ToolStripSeparator());

            var btnCut = new ToolStripButton("切取");
            btnCut.ToolTipText = "切り取り (Ctrl+X)";
            btnCut.Click += Menu_Cut_Click;
            toolStrip.Items.Add(btnCut);

            var btnCopy = new ToolStripButton("コピー");
            btnCopy.ToolTipText = "コピー (Ctrl+C)";
            btnCopy.Click += Menu_Copy_Click;
            toolStrip.Items.Add(btnCopy);

            var btnPaste = new ToolStripButton("貼付");
            btnPaste.ToolTipText = "貼り付け (Ctrl+V)";
            btnPaste.Click += Menu_Paste_Click;
            toolStrip.Items.Add(btnPaste);

            toolStrip.Items.Add(new ToolStripSeparator());

            var btnFind = new ToolStripButton("検索");
            btnFind.ToolTipText = "検索 (Ctrl+F)";
            btnFind.Click += Menu_Find_Click;
            toolStrip.Items.Add(btnFind);

            var btnReplace = new ToolStripButton("置換");
            btnReplace.ToolTipText = "置換 (Ctrl+H)";
            btnReplace.Click += Menu_Replace_Click;
            toolStrip.Items.Add(btnReplace);

            toolStrip.Items.Add(new ToolStripSeparator());

            // 表モード切替ボタン
            btnTableMode = new ToolStripButton("表モード");
            btnTableMode.ToolTipText = "表モード切り替え (Ctrl+T)";
            btnTableMode.CheckOnClick = true;
            btnTableMode.Click += BtnTableMode_Click;
            toolStrip.Items.Add(btnTableMode);

            var btnColumnSearch = new ToolStripButton("列検索");
            btnColumnSearch.ToolTipText = "列の検索・置換 (Ctrl+L)";
            btnColumnSearch.Click += Menu_ColumnSearch_Click;
            toolStrip.Items.Add(btnColumnSearch);

            this.Controls.Add(toolStrip);
        }

        /// <summary>
        /// ステータスストリップ初期化
        /// </summary>
        private void InitializeStatusStrip()
        {
            statusStrip = new StatusStrip();
            statusStrip.Dock = DockStyle.Bottom;

            lblMode = new ToolStripStatusLabel();
            lblMode.AutoSize = false;
            lblMode.Width = 80;
            lblMode.TextAlign = ContentAlignment.MiddleLeft;
            lblMode.Text = "テキスト";
            statusStrip.Items.Add(lblMode);

            statusStrip.Items.Add(new ToolStripSeparator());

            lblPosition = new ToolStripStatusLabel();
            lblPosition.AutoSize = false;
            lblPosition.Width = 120;
            lblPosition.TextAlign = ContentAlignment.MiddleLeft;
            lblPosition.Text = "行: 1  列: 1";
            statusStrip.Items.Add(lblPosition);

            statusStrip.Items.Add(new ToolStripSeparator());

            lblSelection = new ToolStripStatusLabel();
            lblSelection.AutoSize = false;
            lblSelection.Width = 100;
            lblSelection.TextAlign = ContentAlignment.MiddleLeft;
            lblSelection.Text = "選択: 0 文字";
            statusStrip.Items.Add(lblSelection);

            statusStrip.Items.Add(new ToolStripSeparator());

            lblCharCode = new ToolStripStatusLabel();
            lblCharCode.AutoSize = false;
            lblCharCode.Width = 150;
            lblCharCode.TextAlign = ContentAlignment.MiddleLeft;
            lblCharCode.Text = "";
            statusStrip.Items.Add(lblCharCode);

            statusStrip.Items.Add(new ToolStripSeparator());

            lblEncoding = new ToolStripStatusLabel();
            lblEncoding.AutoSize = false;
            lblEncoding.Width = 100;
            lblEncoding.TextAlign = ContentAlignment.MiddleLeft;
            lblEncoding.Text = "UTF-8";
            statusStrip.Items.Add(lblEncoding);

            statusStrip.Items.Add(new ToolStripSeparator());

            lblNewLine = new ToolStripStatusLabel();
            lblNewLine.AutoSize = false;
            lblNewLine.Width = 60;
            lblNewLine.TextAlign = ContentAlignment.MiddleLeft;
            lblNewLine.Text = "CRLF";
            statusStrip.Items.Add(lblNewLine);

            this.Controls.Add(statusStrip);
        }

        /// <summary>
        /// エディタパネル初期化
        /// </summary>
        private void InitializeEditorPanel()
        {
            panelEditor = new Panel();
            panelEditor.Dock = DockStyle.Fill;

            lineNumberPanel = new LineNumberPanel();
            lineNumberPanel.Dock = DockStyle.Left;
            lineNumberPanel.Width = 50;

            textEditor = new SpecialCharRichTextBox();  // 特殊文字表示対応
            textEditor.Dock = DockStyle.Fill;
            textEditor.Font = settings.EditorFont;
            textEditor.WordWrap = false;
            textEditor.AcceptsTab = true;
            textEditor.ScrollBars = RichTextBoxScrollBars.Both;
            textEditor.HideSelection = false;
            textEditor.BorderStyle = BorderStyle.None;
            textEditor.DetectUrls = false;
            textEditor.EnableAutoDragDrop = false; // 自動ドラッグ＆ドロップを無効化（ファイルドロップ用）
            textEditor.ShowSpecialChars = true;    // デフォルトで特殊文字表示ON

            // ドラッグ＆ドロップを有効化（RichTextBox）
            textEditor.AllowDrop = true;
            textEditor.DragEnter += Control_DragEnter;
            textEditor.DragDrop += Control_DragDrop;

            textEditor.TextChanged += TextEditor_TextChanged;
            textEditor.SelectionChanged += TextEditor_SelectionChanged;
            textEditor.VScroll += TextEditor_VScroll;
            textEditor.Resize += TextEditor_Resize;
            textEditor.FontChanged += TextEditor_FontChanged;

            lineNumberPanel.SetTargetTextBox(textEditor);
            lineNumberPanel.SetFont(settings.EditorFont);

            // ドラッグ＆ドロップを有効化（行番号パネル）
            lineNumberPanel.AllowDrop = true;
            lineNumberPanel.DragEnter += Control_DragEnter;
            lineNumberPanel.DragDrop += Control_DragDrop;

            // ドラッグ＆ドロップを有効化（エディタパネル）
            panelEditor.AllowDrop = true;
            panelEditor.DragEnter += Control_DragEnter;
            panelEditor.DragDrop += Control_DragDrop;

            panelEditor.Controls.Add(textEditor);
            panelEditor.Controls.Add(lineNumberPanel);

            this.Controls.Add(panelEditor);
            panelEditor.BringToFront();
        }

        /// <summary>
        /// 表モードパネル初期化
        /// </summary>
        private void InitializeTableModePanel()
        {
            tableModePanel = new TableModePanel();
            tableModePanel.Visible = false;
            tableModePanel.DataChanged += TableModePanel_DataChanged;
            tableModePanel.ColumnSelected += TableModePanel_ColumnSelected;
            tableModePanel.ReParseRequested += TableModePanel_ReParseRequested;
            tableModePanel.SetFont(settings.EditorFont);

            // ドラッグ＆ドロップを有効化（表モードパネル）
            tableModePanel.AllowDrop = true;
            tableModePanel.DragEnter += Control_DragEnter;
            tableModePanel.DragDrop += Control_DragDrop;

            // DataGridViewにもドラッグ＆ドロップを設定
            tableModePanel.Grid.AllowDrop = true;
            tableModePanel.Grid.DragEnter += Control_DragEnter;
            tableModePanel.Grid.DragDrop += Control_DragDrop;

            this.Controls.Add(tableModePanel);
        }

        /// <summary>
        /// 表モードで再解析が要求された時
        /// </summary>
        private void TableModePanel_ReParseRequested(object sender, EventArgs e)
        {
            // 表モードのデータが変更されたことを記録
            settings.IsModified = true;
            UpdateTitle();
        }

        /// <summary>
        /// 表モードのデータ変更時
        /// </summary>
        private void TableModePanel_DataChanged(object sender, EventArgs e)
        {
            settings.IsModified = true;
            UpdateTitle();
        }

        /// <summary>
        /// 表モードで列選択時
        /// </summary>
        private void TableModePanel_ColumnSelected(object sender, ColumnSelectedEventArgs e)
        {
            lblSelection.Text = string.Format("選択列: {0}", e.ColumnName);
        }

        /// <summary>
        /// 表モード切り替え
        /// </summary>
        private void ToggleTableMode()
        {
            if (isTableMode)
            {
                // 表モード → テキストモード
                SyncTableToText();
                tableModePanel.Visible = false;
                panelEditor.Visible = true;
                isTableMode = false;
                lblMode.Text = "テキスト";
            }
            else
            {
                // テキストモード → 表モード
                SyncTextToTable();
                panelEditor.Visible = false;
                tableModePanel.Visible = true;
                tableModePanel.BringToFront();
                isTableMode = true;
                lblMode.Text = "表モード";
            }

            btnTableMode.Checked = isTableMode;
            menuTableMode.Checked = isTableMode;
            UpdateStatusBar();
        }

        /// <summary>
        /// テキストを表に同期
        /// </summary>
        private void SyncTextToTable()
        {
            tableModePanel.LoadFromText(textEditor.Text);
        }

        /// <summary>
        /// 表をテキストに同期
        /// </summary>
        private void SyncTableToText()
        {
            string newText = tableModePanel.ToText();
            if (textEditor.Text != newText)
            {
                textEditor.Text = newText;
            }
        }

        /// <summary>
        /// タイトルバー更新
        /// </summary>
        private void UpdateTitle()
        {
            string fileName = string.IsNullOrEmpty(settings.CurrentFilePath) 
                ? "新規" 
                : Path.GetFileName(settings.CurrentFilePath);
            
            string modified = settings.IsModified ? " *" : "";
            
            this.Text = string.Format("{0}{1} - {2}", fileName, modified, AppName);
        }

        /// <summary>
        /// ステータスバー更新
        /// </summary>
        /// <summary>
        /// ステータスバー更新（即時）
        /// </summary>
        private void UpdateStatusBar()
        {
            UpdateStatusBarInternal();
        }

        /// <summary>
        /// ステータスバー更新（内部実装）
        /// </summary>
        private void UpdateStatusBarInternal()
        {
            if (isTableMode)
            {
                var grid = tableModePanel.Grid;
                int rowCount = grid.Rows.Count;
                int colCount = grid.Columns.Count;
                lblPosition.Text = string.Format("行:{0} 列:{1}", rowCount, colCount);
                
                int selectedColIndex = tableModePanel.GetSelectedColumnIndex();
                if (selectedColIndex >= 0)
                {
                    lblSelection.Text = string.Format("選択列: {0}", grid.Columns[selectedColIndex].HeaderText);
                }
                else
                {
                    lblSelection.Text = string.Format("セル: {0}", grid.SelectedCells.Count);
                }

                lblCharCode.Text = "";  // 表モードでは文字コード表示なし
            }
            else
            {
                if (textEditor == null)
                {
                    return;
                }

                int charIndex = textEditor.SelectionStart;
                int line = textEditor.GetLineFromCharIndex(charIndex);
                int firstCharOfLine = textEditor.GetFirstCharIndexFromLine(line);
                int column = charIndex - firstCharOfLine;

                lblPosition.Text = string.Format("行: {0}  列: {1}", line + 1, column + 1);
                lblSelection.Text = string.Format("選択: {0} 文字", textEditor.SelectionLength);

                // カーソル位置の文字コードを表示
                UpdateCharCodeDisplay(charIndex);
            }

            lblEncoding.Text = settings.GetEncodingDisplayName();
            lblNewLine.Text = settings.GetNewLineDisplayName();
        }

        /// <summary>
        /// カーソル位置の文字コードをステータスバーに表示
        /// </summary>
        private void UpdateCharCodeDisplay(int charIndex)
        {
            if (textEditor == null || string.IsNullOrEmpty(textEditor.Text))
            {
                lblCharCode.Text = "";
                return;
            }

            if (charIndex < 0 || charIndex >= textEditor.Text.Length)
            {
                lblCharCode.Text = "";
                return;
            }

            char c = textEditor.Text[charIndex];
            
            // 現在のエンコーディングでバイト列を取得
            string charStr = c.ToString();
            
            // サロゲートペアの処理
            if (char.IsHighSurrogate(c) && charIndex + 1 < textEditor.Text.Length)
            {
                char lowSurrogate = textEditor.Text[charIndex + 1];
                if (char.IsLowSurrogate(lowSurrogate))
                {
                    charStr = new string(new[] { c, lowSurrogate });
                }
            }

            try
            {
                byte[] bytes = settings.CurrentEncoding.GetBytes(charStr);
                string hexStr = BitConverter.ToString(bytes).Replace("-", " ");
                string charDesc = GetCharDescription(c);
                string displayChar = GetDisplayChar(c);

                // 表示形式: 文字 [XX XX] (説明)
                lblCharCode.Text = string.Format("'{0}' [{1}] {2}", displayChar, hexStr, charDesc);
            }
            catch
            {
                lblCharCode.Text = "";
            }
        }

        /// <summary>
        /// 表示用の文字を取得（制御文字などは変換）
        /// </summary>
        private string GetDisplayChar(char c)
        {
            if (c == '\n') return "LF";
            if (c == '\r') return "CR";
            if (c == '\t') return "TAB";
            if (c == ' ') return "SP";
            if (c == '\u3000') return "全角SP";
            if (char.IsControl(c)) return "CTRL";
            return c.ToString();
        }

        /// <summary>
        /// 文字の説明を取得
        /// </summary>
        private string GetCharDescription(char c)
        {
            if (c == '\n') return "(改行LF)";
            if (c == '\r') return "(復帰CR)";
            if (c == '\t') return "(タブ)";
            if (c == ' ') return "(半角空白)";
            if (c == '\u3000') return "(全角空白)";
            if (char.IsControl(c)) return "(制御文字)";
            if (c >= 0x3040 && c <= 0x309F) return "(ひらがな)";
            if (c >= 0x30A0 && c <= 0x30FF) return "(カタカナ)";
            if (c >= 0x4E00 && c <= 0x9FFF) return "(漢字)";
            if (c >= 0xFF00 && c <= 0xFFEF) return "(全角)";
            return "";
        }

        /// <summary>
        /// 行番号パネルの幅を更新
        /// </summary>
        private void UpdateLineNumberPanelWidth()
        {
            if (lineNumberPanel != null)
            {
                int requiredWidth = lineNumberPanel.CalculateRequiredWidth();
                if (lineNumberPanel.Width != requiredWidth)
                {
                    lineNumberPanel.Width = requiredWidth;
                }
            }
        }

        // イベントハンドラ

        private void TextEditor_TextChanged(object sender, EventArgs e)
        {
            settings.IsModified = true;
            UpdateTitle();
            UpdateLineNumberPanelWidth();
            lineNumberPanel.Invalidate();
        }

        private void TextEditor_SelectionChanged(object sender, EventArgs e)
        {
            RequestStatusUpdate();
        }

        private void TextEditor_VScroll(object sender, EventArgs e)
        {
            lineNumberPanel.Invalidate();
        }

        private void TextEditor_Resize(object sender, EventArgs e)
        {
            lineNumberPanel.Invalidate();
        }

        private void TextEditor_FontChanged(object sender, EventArgs e)
        {
            lineNumberPanel.SetFont(textEditor.Font);
            UpdateLineNumberPanelWidth();
            lineNumberPanel.Invalidate();
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 追加のショートカット処理が必要な場合
        }

        /// <summary>
        /// ドラッグエンター時（共通）
        /// </summary>
        private void Control_DragEnter(object sender, DragEventArgs e)
        {
            // ファイルがドラッグされている場合はコピーカーソルを表示
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// ドラッグドロップ時（共通）
        /// </summary>
        private void Control_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            
            if (files == null || files.Length == 0)
            {
                return;
            }

            // 最初のファイルのみ開く
            string filePath = files[0];

            // ディレクトリはスキップ
            if (Directory.Exists(filePath))
            {
                MessageBox.Show("フォルダは開けません。ファイルをドロップしてください。",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show("ファイルが見つかりません。",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 変更確認
            if (!ConfirmSaveIfModified())
            {
                return;
            }

            // 文字コード自動判定でファイルを開く
            OpenFileWithAutoDetect(filePath);
        }

        /// <summary>
        /// 文字コード自動判定でファイルを開く
        /// </summary>
        private void OpenFileWithAutoDetect(string filePath)
        {
            try
            {
                // 表モードを解除
                if (isTableMode)
                {
                    ToggleTableMode();
                }

                // 文字コード自動判定
                var detection = EncodingDetector.DetectEncoding(filePath);
                
                if (!detection.IsDetected)
                {
                    // 判定できなかった場合はメッセージを表示
                    MessageBox.Show(
                        "文字コードを判定できませんでした。UTF-8として開きます。",
                        "文字コード判定",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                // ファイルを読み込む
                string content;
                using (var reader = new StreamReader(filePath, detection.Encoding, true))
                {
                    content = reader.ReadToEnd();
                    // 実際に使用された文字コードを取得（BOM検出などで変わる可能性）
                    settings.CurrentEncoding = reader.CurrentEncoding;
                }

                // 改行コード検出
                settings.DetectedNewLine = DetectNewLineType(content);
                content = content.Replace("\r\n", "\n").Replace("\r", "\n");

                // テキストを設定
                textEditor.Text = content;
                settings.CurrentFilePath = filePath;
                settings.IsModified = false;

                // CSV/TSVファイルの場合は区切り文字を自動設定
                string ext = Path.GetExtension(filePath).ToLower();
                if (ext == ".csv")
                {
                    tableModePanel.Delimiter = ',';
                }
                else if (ext == ".tsv")
                {
                    tableModePanel.Delimiter = '\t';
                }

                UpdateTitle();
                UpdateStatusBar();
                UpdateLineNumberPanelWidth();
                lineNumberPanel.Invalidate();

                textEditor.SelectionStart = 0;
                textEditor.ScrollToCaret();

                // 検出結果をステータスバーで確認できるようにする
                // （文字コード表示は自動的にUpdateStatusBarで更新される）
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("ファイルを開けませんでした。\n\n{0}", ex.Message),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 表モードの場合はテキストに同期
            if (isTableMode)
            {
                SyncTableToText();
            }

            if (!ConfirmSaveIfModified())
            {
                e.Cancel = true;
                return;
            }

            if (searchForm != null)
            {
                searchForm.Dispose();
                searchForm = null;
            }

            if (columnSearchForm != null)
            {
                columnSearchForm.Dispose();
                columnSearchForm = null;
            }

            base.OnFormClosing(e);
        }

        private bool ConfirmSaveIfModified()
        {
            if (!settings.IsModified)
            {
                return true;
            }

            string fileName = string.IsNullOrEmpty(settings.CurrentFilePath) 
                ? "新規" 
                : Path.GetFileName(settings.CurrentFilePath);

            DialogResult result = MessageBox.Show(
                string.Format("'{0}' への変更を保存しますか？", fileName),
                "保存の確認",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            switch (result)
            {
                case DialogResult.Yes:
                    return SaveFile();
                case DialogResult.No:
                    return true;
                case DialogResult.Cancel:
                default:
                    return false;
            }
        }

        // ファイル操作

        private void NewFile()
        {
            if (!ConfirmSaveIfModified())
            {
                return;
            }

            // 表モードを解除
            if (isTableMode)
            {
                ToggleTableMode();
            }

            textEditor.Clear();
            settings.CurrentFilePath = null;
            settings.IsModified = false;
            settings.CurrentEncoding = new UTF8Encoding(false);
            settings.DetectedNewLine = NewLineType.CRLF;
            
            UpdateTitle();
            UpdateStatusBar();
            lineNumberPanel.Invalidate();
        }

        private void OpenFileWithDialog()
        {
            if (!ConfirmSaveIfModified())
            {
                return;
            }

            Encoding selectedEncoding;
            using (var encodingForm = new EncodingSelectForm(settings.CurrentEncoding))
            {
                if (encodingForm.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }
                selectedEncoding = encodingForm.SelectedEncoding;
            }

            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "テキストファイル (*.txt;*.csv;*.tsv)|*.txt;*.csv;*.tsv|CSVファイル (*.csv)|*.csv|TSVファイル (*.tsv)|*.tsv|すべてのファイル (*.*)|*.*";
                dialog.FilterIndex = 4;
                dialog.Title = "ファイルを開く";

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    OpenFile(dialog.FileName, selectedEncoding);
                }
            }
        }

        private void OpenFile(string filePath, Encoding encoding)
        {
            try
            {
                // 表モードを解除
                if (isTableMode)
                {
                    ToggleTableMode();
                }

                string content;
                using (var reader = new StreamReader(filePath, encoding, true))
                {
                    content = reader.ReadToEnd();
                    settings.CurrentEncoding = reader.CurrentEncoding;
                }

                settings.DetectedNewLine = DetectNewLineType(content);
                content = content.Replace("\r\n", "\n").Replace("\r", "\n");

                textEditor.Text = content;
                settings.CurrentFilePath = filePath;
                settings.IsModified = false;

                // CSV/TSVファイルの場合は区切り文字を自動設定
                string ext = Path.GetExtension(filePath).ToLower();
                if (ext == ".csv")
                {
                    tableModePanel.Delimiter = ',';
                }
                else if (ext == ".tsv")
                {
                    tableModePanel.Delimiter = '\t';
                }

                UpdateTitle();
                UpdateStatusBar();
                UpdateLineNumberPanelWidth();
                lineNumberPanel.Invalidate();

                textEditor.SelectionStart = 0;
                textEditor.ScrollToCaret();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("ファイルを開けませんでした。\n\n{0}", ex.Message),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private NewLineType DetectNewLineType(string content)
        {
            int crlfCount = 0;
            int lfCount = 0;
            int crCount = 0;

            for (int i = 0; i < content.Length; i++)
            {
                if (content[i] == '\r')
                {
                    if (i + 1 < content.Length && content[i + 1] == '\n')
                    {
                        crlfCount++;
                        i++;
                    }
                    else
                    {
                        crCount++;
                    }
                }
                else if (content[i] == '\n')
                {
                    lfCount++;
                }
            }

            if (crlfCount >= lfCount && crlfCount >= crCount)
            {
                return NewLineType.CRLF;
            }
            else if (lfCount >= crCount)
            {
                return NewLineType.LF;
            }
            else
            {
                return NewLineType.CR;
            }
        }

        private bool SaveFile()
        {
            // 表モードの場合はテキストに同期
            if (isTableMode)
            {
                SyncTableToText();
            }

            if (string.IsNullOrEmpty(settings.CurrentFilePath))
            {
                return SaveFileAs();
            }

            return SaveToFile(settings.CurrentFilePath);
        }

        private bool SaveFileAs()
        {
            // 表モードの場合はテキストに同期
            if (isTableMode)
            {
                SyncTableToText();
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "テキストファイル (*.txt)|*.txt|CSVファイル (*.csv)|*.csv|TSVファイル (*.tsv)|*.tsv|すべてのファイル (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.Title = "名前を付けて保存";
                dialog.DefaultExt = "txt";

                if (!string.IsNullOrEmpty(settings.CurrentFilePath))
                {
                    dialog.FileName = Path.GetFileName(settings.CurrentFilePath);
                    dialog.InitialDirectory = Path.GetDirectoryName(settings.CurrentFilePath);
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    return SaveToFile(dialog.FileName);
                }
            }

            return false;
        }

        private bool SaveToFile(string filePath)
        {
            try
            {
                string content = textEditor.Text;
                string newLine = settings.GetNewLineString();
                content = content.Replace("\n", newLine);

                using (var writer = new StreamWriter(filePath, false, settings.CurrentEncoding))
                {
                    writer.Write(content);
                }

                settings.CurrentFilePath = filePath;
                settings.IsModified = false;
                UpdateTitle();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("ファイルを保存できませんでした。\n\n{0}", ex.Message),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        // メニューイベントハンドラ

        private void Menu_New_Click(object sender, EventArgs e) { NewFile(); }
        private void Menu_Open_Click(object sender, EventArgs e) { OpenFileWithDialog(); }
        private void Menu_Save_Click(object sender, EventArgs e) { SaveFile(); }
        private void Menu_SaveAs_Click(object sender, EventArgs e) { SaveFileAs(); }
        private void Menu_Exit_Click(object sender, EventArgs e) { this.Close(); }

        private void Menu_Undo_Click(object sender, EventArgs e)
        {
            if (!isTableMode && textEditor.CanUndo)
            {
                textEditor.Undo();
            }
        }

        private void Menu_Redo_Click(object sender, EventArgs e)
        {
            if (!isTableMode && textEditor.CanRedo)
            {
                textEditor.Redo();
            }
        }

        private void Menu_Cut_Click(object sender, EventArgs e)
        {
            if (!isTableMode)
            {
                textEditor.Cut();
            }
        }

        private void Menu_Copy_Click(object sender, EventArgs e)
        {
            if (isTableMode)
            {
                tableModePanel.CopySelectedCells();
            }
            else
            {
                textEditor.Copy();
            }
        }

        private void Menu_Paste_Click(object sender, EventArgs e)
        {
            if (!isTableMode)
            {
                textEditor.Paste();
            }
        }

        private void Menu_SelectAll_Click(object sender, EventArgs e)
        {
            if (!isTableMode)
            {
                textEditor.SelectAll();
            }
            else
            {
                tableModePanel.Grid.SelectAll();
            }
        }

        private void Menu_TableMode_Click(object sender, EventArgs e)
        {
            ToggleTableMode();
        }

        private void BtnTableMode_Click(object sender, EventArgs e)
        {
            ToggleTableMode();
        }

        private void Menu_Find_Click(object sender, EventArgs e)
        {
            if (isTableMode)
            {
                ShowColumnSearchForm();
            }
            else
            {
                ShowSearchForm(false);
            }
        }

        private void Menu_FindNext_Click(object sender, EventArgs e)
        {
            if (!isTableMode && searchForm != null && !string.IsNullOrEmpty(searchForm.SearchText))
            {
                ShowSearchForm(false);
            }
        }

        private void Menu_FindPrev_Click(object sender, EventArgs e)
        {
            if (!isTableMode)
            {
                ShowSearchForm(false);
            }
        }

        private void Menu_Replace_Click(object sender, EventArgs e)
        {
            if (isTableMode)
            {
                ShowColumnSearchForm();
            }
            else
            {
                ShowSearchForm(true);
            }
        }

        private void Menu_ColumnSearch_Click(object sender, EventArgs e)
        {
            if (!isTableMode)
            {
                // 表モードでない場合は表モードに切り替え
                ToggleTableMode();
            }
            ShowColumnSearchForm();
        }

        private void ShowSearchForm(bool showReplace)
        {
            if (searchForm == null || searchForm.IsDisposed)
            {
                searchForm = new SearchForm(textEditor);
            }

            if (!string.IsNullOrEmpty(textEditor.SelectedText) && 
                !textEditor.SelectedText.Contains("\n"))
            {
                searchForm.SearchText = textEditor.SelectedText;
            }

            if (!searchForm.Visible)
            {
                searchForm.Show(this);
            }
            else
            {
                searchForm.Activate();
            }
        }

        private void ShowColumnSearchForm()
        {
            if (columnSearchForm == null || columnSearchForm.IsDisposed)
            {
                columnSearchForm = new ColumnSearchForm(tableModePanel.Grid);
            }

            if (!columnSearchForm.Visible)
            {
                columnSearchForm.Show(this);
            }
            else
            {
                columnSearchForm.Activate();
            }
        }

        private void Menu_GoToLine_Click(object sender, EventArgs e)
        {
            if (isTableMode)
            {
                return;
            }

            int maxLine = textEditor.Lines.Length;
            if (maxLine < 1) maxLine = 1;

            int currentLine = textEditor.GetLineFromCharIndex(textEditor.SelectionStart) + 1;

            using (var dialog = new GoToLineForm(maxLine, currentLine))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    int lineIndex = dialog.SelectedLineNumber - 1;
                    int charIndex = textEditor.GetFirstCharIndexFromLine(lineIndex);
                    
                    if (charIndex >= 0)
                    {
                        textEditor.SelectionStart = charIndex;
                        textEditor.SelectionLength = 0;
                        textEditor.ScrollToCaret();
                        textEditor.Focus();
                    }
                }
            }
        }

        private void Menu_Font_Click(object sender, EventArgs e)
        {
            using (var dialog = new FontDialog())
            {
                dialog.Font = textEditor.Font;
                dialog.ShowEffects = false;
                dialog.FontMustExist = true;

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    textEditor.Font = dialog.Font;
                    settings.EditorFont = dialog.Font;
                    tableModePanel.SetFont(dialog.Font);
                }
            }
        }

        private void Menu_Encoding_Click(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            if (menuItem != null && menuItem.Tag is Encoding)
            {
                settings.CurrentEncoding = (Encoding)menuItem.Tag;
                settings.IsModified = true;
                UpdateTitle();
                UpdateStatusBar();
            }
        }

        private void Menu_ShowSpecialChars_Click(object sender, EventArgs e)
        {
            // 特殊文字表示の切り替え
            textEditor.ShowSpecialChars = menuShowSpecialChars.Checked;
            textEditor.Invalidate();
        }

        private void Menu_About_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                string.Format("{0}\n\nVersion 1.4.0\n\n山口さん専用テキストエディタ\n\n" +
                    "機能:\n・テキストファイルの編集\n・検索・置換（正規表現対応）\n" +
                    "・行番号表示\n・文字コード切り替え\n・フォント変更\n" +
                    "・表モード（CSV/TSV対応）\n・列単位の検索・置換・コピー\n" +
                    "・ドラッグ＆ドロップ（文字コード自動判定）\n" +
                    "・特殊文字表示（空白・改行の可視化）\n" +
                    "・文字コード表示（ステータスバー）\n" +
                    "・表モードの行絞り込み",
                    AppName),
                "バージョン情報",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (searchForm != null)
                {
                    searchForm.Dispose();
                    searchForm = null;
                }
                if (columnSearchForm != null)
                {
                    columnSearchForm.Dispose();
                    columnSearchForm = null;
                }
                if (statusUpdateTimer != null)
                {
                    statusUpdateTimer.Stop();
                    statusUpdateTimer.Dispose();
                    statusUpdateTimer = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
