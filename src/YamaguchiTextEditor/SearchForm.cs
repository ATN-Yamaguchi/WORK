using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 検索・置換ダイアログ（モードレス）
    /// </summary>
    public class SearchForm : Form
    {
        private Label lblSearch;
        private Label lblReplace;
        private TextBox txtSearch;
        private TextBox txtReplace;
        private CheckBox chkMatchCase;
        private CheckBox chkRegex;
        private Button btnFindNext;
        private Button btnFindPrev;
        private Button btnReplace;
        private Button btnReplaceAll;
        private Button btnClose;

        private RichTextBox _targetTextBox;
        private string _lastSearchText;
        private bool _lastMatchCase;
        private bool _lastUseRegex;

        /// <summary>
        /// 検索文字列プロパティ
        /// </summary>
        public string SearchText
        {
            get { return txtSearch.Text; }
            set { txtSearch.Text = value; }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="targetTextBox">検索対象のRichTextBox</param>
        public SearchForm(RichTextBox targetTextBox)
        {
            _targetTextBox = targetTextBox;
            _lastSearchText = string.Empty;
            _lastMatchCase = false;
            _lastUseRegex = false;

            InitializeComponents();
        }

        /// <summary>
        /// コンポーネント初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Text = "検索と置換";
            this.Size = new Size(420, 220);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.KeyPreview = true;
            this.KeyDown += SearchForm_KeyDown;

            // 検索文字列
            lblSearch = new Label();
            lblSearch.Text = "検索する文字列(&N):";
            lblSearch.Location = new Point(12, 15);
            lblSearch.Size = new Size(120, 20);
            this.Controls.Add(lblSearch);

            txtSearch = new TextBox();
            txtSearch.Location = new Point(135, 12);
            txtSearch.Size = new Size(260, 25);
            txtSearch.TextChanged += TxtSearch_TextChanged;
            this.Controls.Add(txtSearch);

            // 置換文字列
            lblReplace = new Label();
            lblReplace.Text = "置換後の文字列(&P):";
            lblReplace.Location = new Point(12, 45);
            lblReplace.Size = new Size(120, 20);
            this.Controls.Add(lblReplace);

            txtReplace = new TextBox();
            txtReplace.Location = new Point(135, 42);
            txtReplace.Size = new Size(260, 25);
            this.Controls.Add(txtReplace);

            // オプション
            chkMatchCase = new CheckBox();
            chkMatchCase.Text = "大文字と小文字を区別する(&C)";
            chkMatchCase.Location = new Point(12, 75);
            chkMatchCase.Size = new Size(200, 24);
            this.Controls.Add(chkMatchCase);

            chkRegex = new CheckBox();
            chkRegex.Text = "正規表現を使用する(&E)";
            chkRegex.Location = new Point(12, 100);
            chkRegex.Size = new Size(200, 24);
            this.Controls.Add(chkRegex);

            // ボタン
            int buttonY = 135;
            int buttonWidth = 80;
            int buttonHeight = 28;
            int buttonSpacing = 5;

            btnFindNext = new Button();
            btnFindNext.Text = "次を検索(&F)";
            btnFindNext.Location = new Point(12, buttonY);
            btnFindNext.Size = new Size(buttonWidth, buttonHeight);
            btnFindNext.Click += BtnFindNext_Click;
            this.Controls.Add(btnFindNext);

            btnFindPrev = new Button();
            btnFindPrev.Text = "前を検索(&B)";
            btnFindPrev.Location = new Point(12 + buttonWidth + buttonSpacing, buttonY);
            btnFindPrev.Size = new Size(buttonWidth, buttonHeight);
            btnFindPrev.Click += BtnFindPrev_Click;
            this.Controls.Add(btnFindPrev);

            btnReplace = new Button();
            btnReplace.Text = "置換(&R)";
            btnReplace.Location = new Point(12 + (buttonWidth + buttonSpacing) * 2, buttonY);
            btnReplace.Size = new Size(buttonWidth, buttonHeight);
            btnReplace.Click += BtnReplace_Click;
            this.Controls.Add(btnReplace);

            btnReplaceAll = new Button();
            btnReplaceAll.Text = "全て置換(&A)";
            btnReplaceAll.Location = new Point(12 + (buttonWidth + buttonSpacing) * 3, buttonY);
            btnReplaceAll.Size = new Size(buttonWidth, buttonHeight);
            btnReplaceAll.Click += BtnReplaceAll_Click;
            this.Controls.Add(btnReplaceAll);

            btnClose = new Button();
            btnClose.Text = "閉じる";
            btnClose.Location = new Point(315, buttonY + 35);
            btnClose.Size = new Size(80, buttonHeight);
            btnClose.Click += BtnClose_Click;
            this.Controls.Add(btnClose);

            this.CancelButton = btnClose;
            UpdateButtonStates();
        }

        /// <summary>
        /// キーダウンイベント（Escキーで閉じる）
        /// </summary>
        private void SearchForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Hide();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (btnFindNext.Enabled)
                {
                    BtnFindNext_Click(sender, e);
                }
                e.Handled = true;
            }
        }

        /// <summary>
        /// 検索テキスト変更時
        /// </summary>
        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            UpdateButtonStates();
        }

        /// <summary>
        /// ボタンの有効/無効を更新
        /// </summary>
        private void UpdateButtonStates()
        {
            bool hasSearchText = !string.IsNullOrEmpty(txtSearch.Text);
            btnFindNext.Enabled = hasSearchText;
            btnFindPrev.Enabled = hasSearchText;
            btnReplace.Enabled = hasSearchText;
            btnReplaceAll.Enabled = hasSearchText;
        }

        /// <summary>
        /// 次を検索
        /// </summary>
        private void BtnFindNext_Click(object sender, EventArgs e)
        {
            FindText(true);
        }

        /// <summary>
        /// 前を検索
        /// </summary>
        private void BtnFindPrev_Click(object sender, EventArgs e)
        {
            FindText(false);
        }

        /// <summary>
        /// 検索処理
        /// </summary>
        /// <param name="forward">前方検索の場合true</param>
        private void FindText(bool forward)
        {
            if (_targetTextBox == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            string searchText = txtSearch.Text;
            string content = _targetTextBox.Text;
            bool matchCase = chkMatchCase.Checked;
            bool useRegex = chkRegex.Checked;

            // 検索開始位置
            int startIndex;
            if (forward)
            {
                startIndex = _targetTextBox.SelectionStart + _targetTextBox.SelectionLength;
            }
            else
            {
                startIndex = _targetTextBox.SelectionStart - 1;
                if (startIndex < 0)
                {
                    startIndex = content.Length - 1;
                }
            }

            int foundIndex = -1;
            int foundLength = searchText.Length;

            try
            {
                if (useRegex)
                {
                    // 正規表現検索
                    RegexOptions options = RegexOptions.None;
                    if (!matchCase)
                    {
                        options |= RegexOptions.IgnoreCase;
                    }

                    Regex regex = new Regex(searchText, options, TimeSpan.FromSeconds(5));

                    if (forward)
                    {
                        // 前方検索
                        Match match = regex.Match(content, startIndex);
                        if (match.Success)
                        {
                            foundIndex = match.Index;
                            foundLength = match.Length;
                        }
                        else
                        {
                            // 先頭から再検索
                            match = regex.Match(content, 0);
                            if (match.Success && match.Index < _targetTextBox.SelectionStart)
                            {
                                foundIndex = match.Index;
                                foundLength = match.Length;
                            }
                        }
                    }
                    else
                    {
                        // 後方検索（先頭から検索して、開始位置より前の最後のマッチを探す）
                        MatchCollection matches = regex.Matches(content);
                        for (int i = matches.Count - 1; i >= 0; i--)
                        {
                            if (matches[i].Index < startIndex + 1)
                            {
                                foundIndex = matches[i].Index;
                                foundLength = matches[i].Length;
                                break;
                            }
                        }
                        
                        if (foundIndex == -1 && matches.Count > 0)
                        {
                            // 末尾から再検索
                            foundIndex = matches[matches.Count - 1].Index;
                            foundLength = matches[matches.Count - 1].Length;
                        }
                    }
                }
                else
                {
                    // 通常検索
                    StringComparison comparison = matchCase 
                        ? StringComparison.Ordinal 
                        : StringComparison.OrdinalIgnoreCase;

                    if (forward)
                    {
                        // 前方検索
                        if (startIndex < content.Length)
                        {
                            foundIndex = content.IndexOf(searchText, startIndex, comparison);
                        }
                        
                        if (foundIndex == -1)
                        {
                            // 先頭から再検索
                            foundIndex = content.IndexOf(searchText, 0, comparison);
                        }
                    }
                    else
                    {
                        // 後方検索
                        if (startIndex >= 0)
                        {
                            int searchLength = Math.Min(startIndex + 1, content.Length);
                            foundIndex = content.LastIndexOf(searchText, startIndex, searchLength, comparison);
                        }
                        
                        if (foundIndex == -1)
                        {
                            // 末尾から再検索
                            foundIndex = content.LastIndexOf(searchText, content.Length - 1, comparison);
                        }
                    }
                }
            }
            catch (ArgumentException ex)
            {
                // .NET Framework 4.8.1では正規表現エラーはArgumentExceptionとしてスローされる
                MessageBox.Show(
                    "正規表現の構文エラー:\n" + ex.Message, 
                    "エラー", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
                return;
            }
            catch (RegexMatchTimeoutException)
            {
                MessageBox.Show(
                    "検索がタイムアウトしました。検索パターンを見直してください。", 
                    "タイムアウト", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
                return;
            }

            if (foundIndex >= 0)
            {
                _targetTextBox.Select(foundIndex, foundLength);
                _targetTextBox.ScrollToCaret();
                _targetTextBox.Focus();
            }
            else
            {
                MessageBox.Show(
                    "検索文字列が見つかりませんでした。", 
                    "検索", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
            }

            _lastSearchText = searchText;
            _lastMatchCase = matchCase;
            _lastUseRegex = useRegex;
        }

        /// <summary>
        /// 置換
        /// </summary>
        private void BtnReplace_Click(object sender, EventArgs e)
        {
            if (_targetTextBox == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            string searchText = txtSearch.Text;
            string replaceText = txtReplace.Text ?? string.Empty;
            bool matchCase = chkMatchCase.Checked;
            bool useRegex = chkRegex.Checked;

            // 現在の選択範囲が検索文字列と一致するか確認
            string selectedText = _targetTextBox.SelectedText;
            bool isMatch = false;

            try
            {
                if (useRegex)
                {
                    RegexOptions options = RegexOptions.None;
                    if (!matchCase)
                    {
                        options |= RegexOptions.IgnoreCase;
                    }
                    Regex regex = new Regex("^" + searchText + "$", options, TimeSpan.FromSeconds(5));
                    isMatch = regex.IsMatch(selectedText);
                }
                else
                {
                    StringComparison comparison = matchCase 
                        ? StringComparison.Ordinal 
                        : StringComparison.OrdinalIgnoreCase;
                    isMatch = selectedText.Equals(searchText, comparison);
                }
            }
            catch (ArgumentException ex)
            {
                // .NET Framework 4.8.1では正規表現エラーはArgumentExceptionとしてスローされる
                MessageBox.Show(
                    "正規表現の構文エラー:\n" + ex.Message, 
                    "エラー", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
                return;
            }

            if (isMatch)
            {
                // 置換実行
                if (useRegex)
                {
                    try
                    {
                        RegexOptions options = RegexOptions.None;
                        if (!matchCase)
                        {
                            options |= RegexOptions.IgnoreCase;
                        }
                        Regex regex = new Regex(searchText, options, TimeSpan.FromSeconds(5));
                        string replaced = regex.Replace(selectedText, replaceText);
                        _targetTextBox.SelectedText = replaced;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "置換エラー:\n" + ex.Message, 
                            "エラー", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    _targetTextBox.SelectedText = replaceText;
                }
            }

            // 次を検索
            FindText(true);
        }

        /// <summary>
        /// 全て置換
        /// </summary>
        private void BtnReplaceAll_Click(object sender, EventArgs e)
        {
            if (_targetTextBox == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            string searchText = txtSearch.Text;
            string replaceText = txtReplace.Text ?? string.Empty;
            string content = _targetTextBox.Text;
            bool matchCase = chkMatchCase.Checked;
            bool useRegex = chkRegex.Checked;

            int replaceCount = 0;
            string newContent;

            try
            {
                if (useRegex)
                {
                    RegexOptions options = RegexOptions.None;
                    if (!matchCase)
                    {
                        options |= RegexOptions.IgnoreCase;
                    }
                    Regex regex = new Regex(searchText, options, TimeSpan.FromSeconds(30));
                    
                    // 置換件数をカウント
                    replaceCount = regex.Matches(content).Count;
                    
                    // 置換実行
                    newContent = regex.Replace(content, replaceText);
                }
                else
                {
                    // 通常の置換
                    StringComparison comparison = matchCase 
                        ? StringComparison.Ordinal 
                        : StringComparison.OrdinalIgnoreCase;

                    // 置換件数をカウント
                    int index = 0;
                    while ((index = content.IndexOf(searchText, index, comparison)) != -1)
                    {
                        replaceCount++;
                        index += searchText.Length;
                    }

                    // 置換実行
                    if (matchCase)
                    {
                        newContent = content.Replace(searchText, replaceText);
                    }
                    else
                    {
                        // 大文字小文字を区別しない置換
                        newContent = ReplaceIgnoreCase(content, searchText, replaceText);
                    }
                }
            }
            catch (ArgumentException ex)
            {
                // .NET Framework 4.8.1では正規表現エラーはArgumentExceptionとしてスローされる
                MessageBox.Show(
                    "正規表現の構文エラー:\n" + ex.Message, 
                    "エラー", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
                return;
            }
            catch (RegexMatchTimeoutException)
            {
                MessageBox.Show(
                    "置換がタイムアウトしました。検索パターンを見直してください。", 
                    "タイムアウト", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
                return;
            }

            if (replaceCount > 0)
            {
                int currentPos = _targetTextBox.SelectionStart;
                _targetTextBox.Text = newContent;
                
                // カーソル位置を可能な限り維持
                if (currentPos <= newContent.Length)
                {
                    _targetTextBox.SelectionStart = currentPos;
                }
                else
                {
                    _targetTextBox.SelectionStart = newContent.Length;
                }
            }

            MessageBox.Show(
                string.Format("{0} 件を置換しました。", replaceCount), 
                "全て置換", 
                MessageBoxButtons.OK, 
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// 大文字小文字を区別しない置換
        /// </summary>
        private string ReplaceIgnoreCase(string source, string oldValue, string newValue)
        {
            System.Text.StringBuilder result = new System.Text.StringBuilder();
            int index = 0;
            int prevIndex = 0;

            while ((index = source.IndexOf(oldValue, prevIndex, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                result.Append(source.Substring(prevIndex, index - prevIndex));
                result.Append(newValue);
                prevIndex = index + oldValue.Length;
            }

            result.Append(source.Substring(prevIndex));
            return result.ToString();
        }

        /// <summary>
        /// 閉じるボタン
        /// </summary>
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        /// <summary>
        /// フォームを閉じる際に非表示にする（破棄しない）
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        /// <summary>
        /// フォーム表示時に検索テキストボックスにフォーカス
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            txtSearch.Focus();
            txtSearch.SelectAll();
        }

        /// <summary>
        /// フォーム再表示時
        /// </summary>
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (this.Visible)
            {
                txtSearch.Focus();
                txtSearch.SelectAll();
            }
        }
    }
}
