using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 列検索・置換ダイアログ
    /// </summary>
    public class ColumnSearchForm : Form
    {
        private Label lblColumn;
        private ComboBox cboColumn;
        private Label lblSearch;
        private TextBox txtSearch;
        private Label lblReplace;
        private TextBox txtReplace;
        private CheckBox chkMatchCase;
        private CheckBox chkRegex;
        private Button btnFindNext;
        private Button btnReplace;
        private Button btnReplaceAll;
        private Button btnClose;

        private DataGridView _targetGrid;
        private int _lastFoundRow;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ColumnSearchForm(DataGridView targetGrid)
        {
            _targetGrid = targetGrid;
            _lastFoundRow = -1;
            InitializeComponents();
            UpdateColumnList();
        }

        /// <summary>
        /// コンポーネント初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Text = "列の検索と置換";
            this.Size = new Size(400, 280);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.KeyPreview = true;
            this.KeyDown += ColumnSearchForm_KeyDown;

            int y = 15;
            int labelWidth = 100;
            int controlX = 110;
            int controlWidth = 260;
            int rowHeight = 30;

            // 対象列
            lblColumn = new Label();
            lblColumn.Text = "対象列(&L):";
            lblColumn.Location = new Point(12, y + 3);
            lblColumn.Size = new Size(labelWidth, 20);
            this.Controls.Add(lblColumn);

            cboColumn = new ComboBox();
            cboColumn.DropDownStyle = ComboBoxStyle.DropDownList;
            cboColumn.Location = new Point(controlX, y);
            cboColumn.Size = new Size(controlWidth, 25);
            this.Controls.Add(cboColumn);

            y += rowHeight;

            // 検索文字列
            lblSearch = new Label();
            lblSearch.Text = "検索する文字列(&N):";
            lblSearch.Location = new Point(12, y + 3);
            lblSearch.Size = new Size(labelWidth, 20);
            this.Controls.Add(lblSearch);

            txtSearch = new TextBox();
            txtSearch.Location = new Point(controlX, y);
            txtSearch.Size = new Size(controlWidth, 25);
            txtSearch.TextChanged += TxtSearch_TextChanged;
            this.Controls.Add(txtSearch);

            y += rowHeight;

            // 置換文字列
            lblReplace = new Label();
            lblReplace.Text = "置換後の文字列(&P):";
            lblReplace.Location = new Point(12, y + 3);
            lblReplace.Size = new Size(labelWidth, 20);
            this.Controls.Add(lblReplace);

            txtReplace = new TextBox();
            txtReplace.Location = new Point(controlX, y);
            txtReplace.Size = new Size(controlWidth, 25);
            this.Controls.Add(txtReplace);

            y += rowHeight + 5;

            // オプション
            chkMatchCase = new CheckBox();
            chkMatchCase.Text = "大文字と小文字を区別する(&C)";
            chkMatchCase.Location = new Point(12, y);
            chkMatchCase.Size = new Size(200, 24);
            this.Controls.Add(chkMatchCase);

            y += 25;

            chkRegex = new CheckBox();
            chkRegex.Text = "正規表現を使用する(&E)";
            chkRegex.Location = new Point(12, y);
            chkRegex.Size = new Size(200, 24);
            this.Controls.Add(chkRegex);

            y += 35;

            // ボタン
            int buttonWidth = 85;
            int buttonHeight = 28;
            int buttonSpacing = 8;
            int buttonX = 12;

            btnFindNext = new Button();
            btnFindNext.Text = "次を検索(&F)";
            btnFindNext.Location = new Point(buttonX, y);
            btnFindNext.Size = new Size(buttonWidth, buttonHeight);
            btnFindNext.Click += BtnFindNext_Click;
            this.Controls.Add(btnFindNext);

            buttonX += buttonWidth + buttonSpacing;

            btnReplace = new Button();
            btnReplace.Text = "置換(&R)";
            btnReplace.Location = new Point(buttonX, y);
            btnReplace.Size = new Size(buttonWidth, buttonHeight);
            btnReplace.Click += BtnReplace_Click;
            this.Controls.Add(btnReplace);

            buttonX += buttonWidth + buttonSpacing;

            btnReplaceAll = new Button();
            btnReplaceAll.Text = "全て置換(&A)";
            btnReplaceAll.Location = new Point(buttonX, y);
            btnReplaceAll.Size = new Size(buttonWidth, buttonHeight);
            btnReplaceAll.Click += BtnReplaceAll_Click;
            this.Controls.Add(btnReplaceAll);

            buttonX += buttonWidth + buttonSpacing;

            btnClose = new Button();
            btnClose.Text = "閉じる";
            btnClose.Location = new Point(buttonX, y);
            btnClose.Size = new Size(buttonWidth, buttonHeight);
            btnClose.Click += BtnClose_Click;
            this.Controls.Add(btnClose);

            this.CancelButton = btnClose;
            UpdateButtonStates();
        }

        /// <summary>
        /// 列リストを更新
        /// </summary>
        public void UpdateColumnList()
        {
            cboColumn.Items.Clear();
            
            if (_targetGrid == null || _targetGrid.Columns.Count == 0)
            {
                return;
            }

            // 全列オプション
            cboColumn.Items.Add("(すべての列)");

            foreach (DataGridViewColumn col in _targetGrid.Columns)
            {
                cboColumn.Items.Add(col.HeaderText);
            }

            // 選択中の列があればそれを選択
            if (_targetGrid.SelectedCells.Count > 0)
            {
                int selectedColIndex = _targetGrid.SelectedCells[0].ColumnIndex;
                cboColumn.SelectedIndex = selectedColIndex + 1;
            }
            else
            {
                cboColumn.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// キーダウン
        /// </summary>
        private void ColumnSearchForm_KeyDown(object sender, KeyEventArgs e)
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
            _lastFoundRow = -1; // 検索テキスト変更時はリセット
        }

        /// <summary>
        /// ボタン状態更新
        /// </summary>
        private void UpdateButtonStates()
        {
            bool hasSearchText = !string.IsNullOrEmpty(txtSearch.Text);
            bool hasColumn = cboColumn.Items.Count > 0;
            
            btnFindNext.Enabled = hasSearchText && hasColumn;
            btnReplace.Enabled = hasSearchText && hasColumn;
            btnReplaceAll.Enabled = hasSearchText && hasColumn;
        }

        /// <summary>
        /// 次を検索
        /// </summary>
        private void BtnFindNext_Click(object sender, EventArgs e)
        {
            if (_targetGrid == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            string searchText = txtSearch.Text;
            bool matchCase = chkMatchCase.Checked;
            bool useRegex = chkRegex.Checked;
            int targetColIndex = cboColumn.SelectedIndex - 1; // -1 = すべての列

            int startRow = _lastFoundRow + 1;
            if (startRow >= _targetGrid.Rows.Count)
            {
                startRow = 0;
            }

            try
            {
                Regex regex = null;
                if (useRegex)
                {
                    RegexOptions options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                    regex = new Regex(searchText, options, TimeSpan.FromSeconds(5));
                }

                // 検索開始位置から末尾まで
                int foundRow = FindInRange(startRow, _targetGrid.Rows.Count - 1, targetColIndex, searchText, matchCase, regex);
                
                // 見つからなければ先頭から検索開始位置まで
                if (foundRow == -1 && startRow > 0)
                {
                    foundRow = FindInRange(0, startRow - 1, targetColIndex, searchText, matchCase, regex);
                }

                if (foundRow >= 0)
                {
                    _lastFoundRow = foundRow;
                }
                else
                {
                    MessageBox.Show("検索文字列が見つかりませんでした。", "検索",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("正規表現の構文エラー:\n" + ex.Message, "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 範囲内を検索
        /// </summary>
        private int FindInRange(int startRow, int endRow, int targetColIndex, string searchText, bool matchCase, Regex regex)
        {
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            for (int row = startRow; row <= endRow; row++)
            {
                int colStart = targetColIndex < 0 ? 0 : targetColIndex;
                int colEnd = targetColIndex < 0 ? _targetGrid.Columns.Count - 1 : targetColIndex;

                for (int col = colStart; col <= colEnd; col++)
                {
                    var cell = _targetGrid.Rows[row].Cells[col];
                    string cellValue = cell.Value?.ToString() ?? string.Empty;

                    bool found = false;
                    if (regex != null)
                    {
                        found = regex.IsMatch(cellValue);
                    }
                    else
                    {
                        found = cellValue.IndexOf(searchText, comparison) >= 0;
                    }

                    if (found)
                    {
                        // セルを選択してスクロール
                        _targetGrid.ClearSelection();
                        _targetGrid.CurrentCell = cell;
                        cell.Selected = true;
                        return row;
                    }
                }
            }

            return -1;
        }

        /// <summary>
        /// 置換
        /// </summary>
        private void BtnReplace_Click(object sender, EventArgs e)
        {
            if (_targetGrid == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            // 現在選択中のセルが検索文字列を含む場合、置換を実行
            if (_targetGrid.CurrentCell != null)
            {
                string cellValue = _targetGrid.CurrentCell.Value?.ToString() ?? string.Empty;
                string searchText = txtSearch.Text;
                string replaceText = txtReplace.Text ?? string.Empty;
                bool matchCase = chkMatchCase.Checked;
                bool useRegex = chkRegex.Checked;

                try
                {
                    bool contains = false;
                    string newValue;

                    if (useRegex)
                    {
                        RegexOptions options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                        Regex regex = new Regex(searchText, options, TimeSpan.FromSeconds(5));
                        contains = regex.IsMatch(cellValue);
                        newValue = regex.Replace(cellValue, replaceText);
                    }
                    else
                    {
                        StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                        contains = cellValue.IndexOf(searchText, comparison) >= 0;
                        
                        if (matchCase)
                        {
                            newValue = cellValue.Replace(searchText, replaceText);
                        }
                        else
                        {
                            newValue = ReplaceIgnoreCase(cellValue, searchText, replaceText);
                        }
                    }

                    if (contains)
                    {
                        _targetGrid.CurrentCell.Value = newValue;
                    }
                }
                catch (ArgumentException ex)
                {
                    MessageBox.Show("正規表現の構文エラー:\n" + ex.Message, "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 次を検索
            BtnFindNext_Click(sender, e);
        }

        /// <summary>
        /// 全て置換
        /// </summary>
        private void BtnReplaceAll_Click(object sender, EventArgs e)
        {
            if (_targetGrid == null || string.IsNullOrEmpty(txtSearch.Text))
            {
                return;
            }

            string searchText = txtSearch.Text;
            string replaceText = txtReplace.Text ?? string.Empty;
            bool matchCase = chkMatchCase.Checked;
            bool useRegex = chkRegex.Checked;
            int targetColIndex = cboColumn.SelectedIndex - 1;

            int replaceCount = 0;

            try
            {
                Regex regex = null;
                if (useRegex)
                {
                    RegexOptions options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                    regex = new Regex(searchText, options, TimeSpan.FromSeconds(30));
                }

                StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                int colStart = targetColIndex < 0 ? 0 : targetColIndex;
                int colEnd = targetColIndex < 0 ? _targetGrid.Columns.Count - 1 : targetColIndex;

                for (int row = 0; row < _targetGrid.Rows.Count; row++)
                {
                    for (int col = colStart; col <= colEnd; col++)
                    {
                        var cell = _targetGrid.Rows[row].Cells[col];
                        string cellValue = cell.Value?.ToString() ?? string.Empty;

                        bool contains = false;
                        string newValue;

                        if (regex != null)
                        {
                            contains = regex.IsMatch(cellValue);
                            newValue = regex.Replace(cellValue, replaceText);
                        }
                        else
                        {
                            contains = cellValue.IndexOf(searchText, comparison) >= 0;
                            
                            if (matchCase)
                            {
                                newValue = cellValue.Replace(searchText, replaceText);
                            }
                            else
                            {
                                newValue = ReplaceIgnoreCase(cellValue, searchText, replaceText);
                            }
                        }

                        if (contains)
                        {
                            cell.Value = newValue;
                            replaceCount++;
                        }
                    }
                }

                MessageBox.Show(string.Format("{0} 件のセルを置換しました。", replaceCount), "全て置換",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("正規表現の構文エラー:\n" + ex.Message, "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 大文字小文字を区別しない置換
        /// </summary>
        private string ReplaceIgnoreCase(string source, string oldValue, string newValue)
        {
            var sb = new System.Text.StringBuilder();
            int index = 0;
            int prevIndex = 0;

            while ((index = source.IndexOf(oldValue, prevIndex, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                sb.Append(source.Substring(prevIndex, index - prevIndex));
                sb.Append(newValue);
                prevIndex = index + oldValue.Length;
            }

            sb.Append(source.Substring(prevIndex));
            return sb.ToString();
        }

        /// <summary>
        /// 閉じる
        /// </summary>
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        /// <summary>
        /// フォームクローズ時は非表示に
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
        /// 表示時
        /// </summary>
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (this.Visible)
            {
                UpdateColumnList();
                txtSearch.Focus();
                txtSearch.SelectAll();
            }
        }
    }
}
