using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 絞り込み条件の種類
    /// </summary>
    public enum FilterType
    {
        None,           // フィルタなし
        Equals,         // 完全一致
        Contains,       // 含む
        StartsWith,     // で始まる
        EndsWith,       // で終わる
        LengthAtLeast,  // 文字数以上
        LengthAtMost    // 文字数以下
    }

    /// <summary>
    /// 表モード用パネル
    /// </summary>
    public class TableModePanel : Panel
    {
        // コントロール - 設定パネル
        private Panel settingsPanel;
        private Label lblDelimiter;
        private ComboBox cboDelimiter;
        private Label lblQuote;
        private ComboBox cboQuote;
        private CheckBox chkFirstRowHeader;
        private Button btnApply;

        // コントロール - 絞り込みパネル
        private Panel filterPanel;
        private Label lblFilterColumn;
        private ComboBox cboFilterColumn;
        private Label lblFilterType;
        private ComboBox cboFilterType;
        private Label lblFilterValue;
        private TextBox txtFilterValue;
        private Button btnFilter;
        private Button btnClearFilter;
        private Label lblFilterStatus;

        // DataGridView
        private DataGridView dataGrid;

        // パーサー
        private CsvParser parser;

        // データ
        private List<List<string>> tableData;           // 全データ
        private List<int> filteredRowIndices;           // フィルタ後の行インデックス（元の行番号）
        private bool isFiltered;                        // フィルタ中かどうか

        // 元のテキスト（再解析用）
        private string originalText;

        // イベント
        public event EventHandler DataChanged;
        public event EventHandler<ColumnSelectedEventArgs> ColumnSelected;
        public event EventHandler ReParseRequested;

        /// <summary>
        /// 区切り文字
        /// </summary>
        public char Delimiter
        {
            get { return parser.Delimiter; }
            set
            {
                parser.Delimiter = value;
                UpdateDelimiterCombo();
            }
        }

        /// <summary>
        /// 囲み文字を使用するか
        /// </summary>
        public bool UseQuotes
        {
            get { return parser.UseQuotes; }
            set
            {
                parser.UseQuotes = value;
                UpdateQuoteCombo();
            }
        }

        /// <summary>
        /// 1行目をヘッダーとして扱うか
        /// </summary>
        public bool FirstRowAsHeader
        {
            get { return chkFirstRowHeader.Checked; }
            set { chkFirstRowHeader.Checked = value; }
        }

        /// <summary>
        /// DataGridView（外部アクセス用）
        /// </summary>
        public DataGridView Grid
        {
            get { return dataGrid; }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public TableModePanel()
        {
            parser = new CsvParser();
            tableData = new List<List<string>>();
            filteredRowIndices = new List<int>();
            isFiltered = false;
            
            InitializeComponents();
        }

        /// <summary>
        /// コンポーネント初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Dock = DockStyle.Fill;

            // 設定パネル
            settingsPanel = new Panel();
            settingsPanel.Dock = DockStyle.Top;
            settingsPanel.Height = 35;
            settingsPanel.BackColor = Color.FromArgb(240, 240, 240);
            settingsPanel.Padding = new Padding(5);

            int x = 10;
            int y = 6;

            lblDelimiter = new Label();
            lblDelimiter.Text = "区切り:";
            lblDelimiter.Location = new Point(x, y + 3);
            lblDelimiter.AutoSize = true;
            settingsPanel.Controls.Add(lblDelimiter);

            x += 45;

            cboDelimiter = new ComboBox();
            cboDelimiter.DropDownStyle = ComboBoxStyle.DropDownList;
            cboDelimiter.Location = new Point(x, y);
            cboDelimiter.Size = new Size(80, 25);
            cboDelimiter.Items.Add("カンマ (,)");
            cboDelimiter.Items.Add("タブ");
            cboDelimiter.SelectedIndex = 0;
            cboDelimiter.SelectedIndexChanged += CboDelimiter_SelectedIndexChanged;
            settingsPanel.Controls.Add(cboDelimiter);

            x += 95;

            lblQuote = new Label();
            lblQuote.Text = "囲み:";
            lblQuote.Location = new Point(x, y + 3);
            lblQuote.AutoSize = true;
            settingsPanel.Controls.Add(lblQuote);

            x += 40;

            cboQuote = new ComboBox();
            cboQuote.DropDownStyle = ComboBoxStyle.DropDownList;
            cboQuote.Location = new Point(x, y);
            cboQuote.Size = new Size(100, 25);
            cboQuote.Items.Add("\" あり");
            cboQuote.Items.Add("なし");
            cboQuote.SelectedIndex = 0;
            cboQuote.SelectedIndexChanged += CboQuote_SelectedIndexChanged;
            settingsPanel.Controls.Add(cboQuote);

            x += 115;

            chkFirstRowHeader = new CheckBox();
            chkFirstRowHeader.Text = "1行目をヘッダーにする";
            chkFirstRowHeader.Location = new Point(x, y + 2);
            chkFirstRowHeader.AutoSize = true;
            chkFirstRowHeader.CheckedChanged += ChkFirstRowHeader_CheckedChanged;
            settingsPanel.Controls.Add(chkFirstRowHeader);

            x += 160;

            btnApply = new Button();
            btnApply.Text = "再解析";
            btnApply.Location = new Point(x, y - 1);
            btnApply.Size = new Size(70, 25);
            btnApply.Click += BtnApply_Click;
            settingsPanel.Controls.Add(btnApply);

            // 絞り込みパネル
            filterPanel = new Panel();
            filterPanel.Dock = DockStyle.Top;
            filterPanel.Height = 35;
            filterPanel.BackColor = Color.FromArgb(230, 240, 250);
            filterPanel.Padding = new Padding(5);

            x = 10;
            y = 6;

            lblFilterColumn = new Label();
            lblFilterColumn.Text = "列:";
            lblFilterColumn.Location = new Point(x, y + 3);
            lblFilterColumn.AutoSize = true;
            filterPanel.Controls.Add(lblFilterColumn);

            x += 25;

            cboFilterColumn = new ComboBox();
            cboFilterColumn.DropDownStyle = ComboBoxStyle.DropDownList;
            cboFilterColumn.Location = new Point(x, y);
            cboFilterColumn.Size = new Size(100, 25);
            filterPanel.Controls.Add(cboFilterColumn);

            x += 110;

            lblFilterType = new Label();
            lblFilterType.Text = "条件:";
            lblFilterType.Location = new Point(x, y + 3);
            lblFilterType.AutoSize = true;
            filterPanel.Controls.Add(lblFilterType);

            x += 35;

            cboFilterType = new ComboBox();
            cboFilterType.DropDownStyle = ComboBoxStyle.DropDownList;
            cboFilterType.Location = new Point(x, y);
            cboFilterType.Size = new Size(120, 25);
            cboFilterType.Items.Add("完全一致");
            cboFilterType.Items.Add("を含む");
            cboFilterType.Items.Add("で始まる");
            cboFilterType.Items.Add("で終わる");
            cboFilterType.Items.Add("文字数以上");
            cboFilterType.Items.Add("文字数以下");
            cboFilterType.SelectedIndex = 1;
            filterPanel.Controls.Add(cboFilterType);

            x += 130;

            lblFilterValue = new Label();
            lblFilterValue.Text = "値:";
            lblFilterValue.Location = new Point(x, y + 3);
            lblFilterValue.AutoSize = true;
            filterPanel.Controls.Add(lblFilterValue);

            x += 25;

            txtFilterValue = new TextBox();
            txtFilterValue.Location = new Point(x, y);
            txtFilterValue.Size = new Size(150, 25);
            txtFilterValue.KeyDown += TxtFilterValue_KeyDown;
            filterPanel.Controls.Add(txtFilterValue);

            x += 160;

            btnFilter = new Button();
            btnFilter.Text = "絞り込み";
            btnFilter.Location = new Point(x, y - 1);
            btnFilter.Size = new Size(75, 25);
            btnFilter.Click += BtnFilter_Click;
            filterPanel.Controls.Add(btnFilter);

            x += 85;

            btnClearFilter = new Button();
            btnClearFilter.Text = "解除";
            btnClearFilter.Location = new Point(x, y - 1);
            btnClearFilter.Size = new Size(50, 25);
            btnClearFilter.Click += BtnClearFilter_Click;
            filterPanel.Controls.Add(btnClearFilter);

            x += 60;

            lblFilterStatus = new Label();
            lblFilterStatus.Text = "";
            lblFilterStatus.Location = new Point(x, y + 3);
            lblFilterStatus.AutoSize = true;
            lblFilterStatus.ForeColor = Color.Blue;
            filterPanel.Controls.Add(lblFilterStatus);

            // DataGridView
            dataGrid = new DataGridView();
            dataGrid.Dock = DockStyle.Fill;
            dataGrid.AllowUserToAddRows = false;
            dataGrid.AllowUserToDeleteRows = false;
            dataGrid.AllowUserToResizeRows = false;
            dataGrid.RowHeadersVisible = true;
            dataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGrid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGrid.MultiSelect = true;
            dataGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGrid.DefaultCellStyle.Font = new Font("MS ゴシック", 10);
            dataGrid.ColumnHeadersDefaultCellStyle.Font = new Font("MS ゴシック", 10, FontStyle.Bold);
            
            // イベント
            dataGrid.CellValueChanged += DataGrid_CellValueChanged;
            dataGrid.ColumnHeaderMouseClick += DataGrid_ColumnHeaderMouseClick;
            dataGrid.KeyDown += DataGrid_KeyDown;
            dataGrid.SelectionChanged += DataGrid_SelectionChanged;

            // 行番号表示
            dataGrid.RowPostPaint += DataGrid_RowPostPaint;

            // パネルに追加（順序が重要：下から上に追加）
            this.Controls.Add(dataGrid);
            this.Controls.Add(filterPanel);
            this.Controls.Add(settingsPanel);
        }

        /// <summary>
        /// 行番号を描画（元の行番号を表示）
        /// </summary>
        private void DataGrid_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            
            // 元の行番号を取得
            int originalRowNumber;
            if (isFiltered && e.RowIndex < filteredRowIndices.Count)
            {
                originalRowNumber = filteredRowIndices[e.RowIndex] + 1;
            }
            else
            {
                originalRowNumber = e.RowIndex + 1;
            }

            // ヘッダー行を考慮
            if (chkFirstRowHeader.Checked)
            {
                originalRowNumber++;
            }

            string rowNumber = originalRowNumber.ToString();
            
            var headerBounds = new Rectangle(
                e.RowBounds.Left, 
                e.RowBounds.Top, 
                grid.RowHeadersWidth, 
                e.RowBounds.Height);

            using (var brush = new SolidBrush(grid.RowHeadersDefaultCellStyle.ForeColor))
            using (var format = new StringFormat())
            {
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(rowNumber, grid.DefaultCellStyle.Font, brush, headerBounds, format);
            }
        }

        /// <summary>
        /// フィルタ値入力時のキーダウン（Enterで絞り込み実行）
        /// </summary>
        private void TxtFilterValue_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnFilter_Click(sender, e);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// 絞り込みボタンクリック
        /// </summary>
        private void BtnFilter_Click(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        /// <summary>
        /// 絞り込み解除ボタンクリック
        /// </summary>
        private void BtnClearFilter_Click(object sender, EventArgs e)
        {
            ClearFilter();
        }

        /// <summary>
        /// 絞り込みを適用
        /// </summary>
        private void ApplyFilter()
        {
            if (tableData == null || tableData.Count == 0)
            {
                return;
            }

            string filterValue = txtFilterValue.Text;
            int filterColumnIndex = cboFilterColumn.SelectedIndex;
            FilterType filterType = GetSelectedFilterType();

            // フィルタ条件がない場合
            if (string.IsNullOrEmpty(filterValue) && filterType != FilterType.LengthAtLeast && filterType != FilterType.LengthAtMost)
            {
                ClearFilter();
                return;
            }

            filteredRowIndices.Clear();
            int dataStartRow = chkFirstRowHeader.Checked ? 1 : 0;

            // 各行をチェック
            for (int row = dataStartRow; row < tableData.Count; row++)
            {
                bool matches = false;

                if (filterColumnIndex == 0) // すべての列
                {
                    for (int col = 0; col < tableData[row].Count; col++)
                    {
                        if (CheckFilterCondition(tableData[row][col], filterValue, filterType))
                        {
                            matches = true;
                            break;
                        }
                    }
                }
                else
                {
                    // 特定の列
                    int colIndex = filterColumnIndex - 1;
                    if (colIndex < tableData[row].Count)
                    {
                        matches = CheckFilterCondition(tableData[row][colIndex], filterValue, filterType);
                    }
                }

                if (matches)
                {
                    filteredRowIndices.Add(row - dataStartRow); // 元の行インデックス（データ行としての）
                }
            }

            isFiltered = true;
            RefreshGridWithFilter();

            // ステータス更新
            int totalRows = tableData.Count - dataStartRow;
            lblFilterStatus.Text = string.Format("絞り込み: {0}/{1}行", filteredRowIndices.Count, totalRows);
        }

        /// <summary>
        /// フィルタ条件をチェック
        /// </summary>
        private bool CheckFilterCondition(string cellValue, string filterValue, FilterType filterType)
        {
            if (cellValue == null) cellValue = string.Empty;

            switch (filterType)
            {
                case FilterType.Equals:
                    return cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);

                case FilterType.Contains:
                    return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;

                case FilterType.StartsWith:
                    return cellValue.StartsWith(filterValue, StringComparison.OrdinalIgnoreCase);

                case FilterType.EndsWith:
                    return cellValue.EndsWith(filterValue, StringComparison.OrdinalIgnoreCase);

                case FilterType.LengthAtLeast:
                    int minLength;
                    if (int.TryParse(filterValue, out minLength))
                    {
                        return cellValue.Length >= minLength;
                    }
                    return false;

                case FilterType.LengthAtMost:
                    int maxLength;
                    if (int.TryParse(filterValue, out maxLength))
                    {
                        return cellValue.Length <= maxLength;
                    }
                    return false;

                default:
                    return true;
            }
        }

        /// <summary>
        /// 選択されたフィルタタイプを取得
        /// </summary>
        private FilterType GetSelectedFilterType()
        {
            switch (cboFilterType.SelectedIndex)
            {
                case 0: return FilterType.Equals;
                case 1: return FilterType.Contains;
                case 2: return FilterType.StartsWith;
                case 3: return FilterType.EndsWith;
                case 4: return FilterType.LengthAtLeast;
                case 5: return FilterType.LengthAtMost;
                default: return FilterType.Contains;
            }
        }

        /// <summary>
        /// 絞り込みを解除
        /// </summary>
        public void ClearFilter()
        {
            isFiltered = false;
            filteredRowIndices.Clear();
            txtFilterValue.Text = "";
            lblFilterStatus.Text = "";
            RefreshGrid();
        }

        /// <summary>
        /// フィルタ適用してグリッドを更新
        /// </summary>
        private void RefreshGridWithFilter()
        {
            dataGrid.Columns.Clear();
            dataGrid.Rows.Clear();

            if (tableData == null || tableData.Count == 0)
            {
                return;
            }

            int colCount = CsvParser.GetMaxColumnCount(tableData);
            int dataStartRow = chkFirstRowHeader.Checked ? 1 : 0;

            // 列を作成
            if (chkFirstRowHeader.Checked && tableData.Count > 0)
            {
                for (int col = 0; col < colCount; col++)
                {
                    string headerText = col < tableData[0].Count ? tableData[0][col] : string.Empty;
                    if (string.IsNullOrEmpty(headerText))
                    {
                        headerText = GetColumnName(col);
                    }
                    dataGrid.Columns.Add("col" + col, headerText);
                    dataGrid.Columns[col].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dataGrid.Columns[col].Width = 100;
                }
            }
            else
            {
                for (int col = 0; col < colCount; col++)
                {
                    dataGrid.Columns.Add("col" + col, GetColumnName(col));
                    dataGrid.Columns[col].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dataGrid.Columns[col].Width = 100;
                }
            }

            // フィルタされた行のみ追加
            foreach (int rowIndex in filteredRowIndices)
            {
                int actualRow = rowIndex + dataStartRow;
                if (actualRow < tableData.Count)
                {
                    var rowData = new string[colCount];
                    for (int col = 0; col < colCount; col++)
                    {
                        rowData[col] = col < tableData[actualRow].Count ? tableData[actualRow][col] : string.Empty;
                    }
                    dataGrid.Rows.Add(rowData);
                }
            }

            // 行ヘッダー幅を調整
            dataGrid.RowHeadersWidth = 60;

            // 列リストを更新
            UpdateFilterColumnList();
        }

        /// <summary>
        /// 区切り文字変更
        /// </summary>
        private void CboDelimiter_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cboDelimiter.SelectedIndex)
            {
                case 0:
                    parser.Delimiter = ',';
                    break;
                case 1:
                    parser.Delimiter = '\t';
                    break;
            }
        }

        /// <summary>
        /// 囲み文字変更
        /// </summary>
        private void CboQuote_SelectedIndexChanged(object sender, EventArgs e)
        {
            parser.UseQuotes = (cboQuote.SelectedIndex == 0);
        }

        /// <summary>
        /// ヘッダー設定変更
        /// </summary>
        private void ChkFirstRowHeader_CheckedChanged(object sender, EventArgs e)
        {
            ClearFilter();
            RefreshGrid();
        }

        /// <summary>
        /// 再解析ボタン
        /// </summary>
        private void BtnApply_Click(object sender, EventArgs e)
        {
            ClearFilter();

            if (!string.IsNullOrEmpty(originalText))
            {
                tableData = parser.Parse(originalText);
                CsvParser.NormalizeColumnCount(tableData);
                RefreshGrid();
            }
            
            ReParseRequested?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// セル値変更時
        /// </summary>
        private void DataGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            // tableDataを更新
            int dataRowIndex;
            if (isFiltered)
            {
                if (e.RowIndex < filteredRowIndices.Count)
                {
                    dataRowIndex = filteredRowIndices[e.RowIndex] + (chkFirstRowHeader.Checked ? 1 : 0);
                }
                else
                {
                    return;
                }
            }
            else
            {
                dataRowIndex = chkFirstRowHeader.Checked ? e.RowIndex + 1 : e.RowIndex;
            }
            
            if (dataRowIndex < tableData.Count && e.ColumnIndex < tableData[dataRowIndex].Count)
            {
                tableData[dataRowIndex][e.ColumnIndex] = dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
                OnDataChanged();
            }
        }

        /// <summary>
        /// 列ヘッダークリック
        /// </summary>
        private void DataGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            SelectColumn(e.ColumnIndex);
            
            // フィルタの列も連動して選択（+1はすべての列が0番目のため）
            if (e.ColumnIndex + 1 < cboFilterColumn.Items.Count)
            {
                cboFilterColumn.SelectedIndex = e.ColumnIndex + 1;
            }
            
            var args = new ColumnSelectedEventArgs(e.ColumnIndex, dataGrid.Columns[e.ColumnIndex].HeaderText);
            ColumnSelected?.Invoke(this, args);
        }

        /// <summary>
        /// セル選択変更時
        /// </summary>
        private void DataGrid_SelectionChanged(object sender, EventArgs e)
        {
            // 選択セルの列をフィルタUIに反映
            if (dataGrid.SelectedCells.Count > 0)
            {
                int colIndex = dataGrid.SelectedCells[0].ColumnIndex;
                if (colIndex + 1 < cboFilterColumn.Items.Count)
                {
                    cboFilterColumn.SelectedIndex = colIndex + 1;
                }
            }
        }

        /// <summary>
        /// 列を選択
        /// </summary>
        public void SelectColumn(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= dataGrid.Columns.Count)
            {
                return;
            }

            dataGrid.ClearSelection();
            
            for (int row = 0; row < dataGrid.Rows.Count; row++)
            {
                dataGrid.Rows[row].Cells[columnIndex].Selected = true;
            }
        }

        /// <summary>
        /// キーダウン
        /// </summary>
        private void DataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedCells();
                e.Handled = true;
            }
        }

        /// <summary>
        /// 選択セルをコピー
        /// </summary>
        public void CopySelectedCells()
        {
            if (dataGrid.SelectedCells.Count == 0)
            {
                return;
            }

            int minRow = int.MaxValue, maxRow = int.MinValue;
            int minCol = int.MaxValue, maxCol = int.MinValue;

            foreach (DataGridViewCell cell in dataGrid.SelectedCells)
            {
                if (cell.RowIndex < minRow) minRow = cell.RowIndex;
                if (cell.RowIndex > maxRow) maxRow = cell.RowIndex;
                if (cell.ColumnIndex < minCol) minCol = cell.ColumnIndex;
                if (cell.ColumnIndex > maxCol) maxCol = cell.ColumnIndex;
            }

            var selectedMap = new bool[maxRow - minRow + 1, maxCol - minCol + 1];
            foreach (DataGridViewCell cell in dataGrid.SelectedCells)
            {
                selectedMap[cell.RowIndex - minRow, cell.ColumnIndex - minCol] = true;
            }

            var sb = new StringBuilder();
            for (int row = 0; row <= maxRow - minRow; row++)
            {
                var rowValues = new List<string>();
                for (int col = 0; col <= maxCol - minCol; col++)
                {
                    if (selectedMap[row, col])
                    {
                        var cell = dataGrid.Rows[row + minRow].Cells[col + minCol];
                        rowValues.Add(cell.Value?.ToString() ?? string.Empty);
                    }
                    else
                    {
                        rowValues.Add(string.Empty);
                    }
                }

                if (maxCol == minCol)
                {
                    sb.AppendLine(rowValues[0]);
                }
                else
                {
                    sb.AppendLine(string.Join("\t", rowValues));
                }
            }

            if (sb.Length > 0)
            {
                Clipboard.SetText(sb.ToString().TrimEnd('\r', '\n'));
            }
        }

        /// <summary>
        /// テキストからロード
        /// </summary>
        public void LoadFromText(string text)
        {
            originalText = text;
            tableData = parser.Parse(text);
            CsvParser.NormalizeColumnCount(tableData);
            ClearFilter();
            RefreshGrid();
        }

        /// <summary>
        /// テキストに変換
        /// </summary>
        public string ToText()
        {
            SyncGridToData();
            return parser.ToText(tableData);
        }

        /// <summary>
        /// DataGridViewの内容をtableDataに同期
        /// </summary>
        private void SyncGridToData()
        {
            // フィルタ中は同期しない（既に個別のセル編集で同期済み）
            if (isFiltered)
            {
                return;
            }

            if (chkFirstRowHeader.Checked && tableData.Count > 0)
            {
                for (int col = 0; col < dataGrid.Columns.Count && col < tableData[0].Count; col++)
                {
                    tableData[0][col] = dataGrid.Columns[col].HeaderText;
                }

                for (int row = 0; row < dataGrid.Rows.Count; row++)
                {
                    int dataRow = row + 1;
                    if (dataRow < tableData.Count)
                    {
                        for (int col = 0; col < dataGrid.Columns.Count && col < tableData[dataRow].Count; col++)
                        {
                            tableData[dataRow][col] = dataGrid.Rows[row].Cells[col].Value?.ToString() ?? string.Empty;
                        }
                    }
                }
            }
            else
            {
                for (int row = 0; row < dataGrid.Rows.Count && row < tableData.Count; row++)
                {
                    for (int col = 0; col < dataGrid.Columns.Count && col < tableData[row].Count; col++)
                    {
                        tableData[row][col] = dataGrid.Rows[row].Cells[col].Value?.ToString() ?? string.Empty;
                    }
                }
            }
        }

        /// <summary>
        /// グリッドを再表示
        /// </summary>
        private void RefreshGrid()
        {
            dataGrid.Columns.Clear();
            dataGrid.Rows.Clear();

            if (tableData == null || tableData.Count == 0)
            {
                return;
            }

            int colCount = CsvParser.GetMaxColumnCount(tableData);
            int dataStartRow = 0;

            if (chkFirstRowHeader.Checked && tableData.Count > 0)
            {
                for (int col = 0; col < colCount; col++)
                {
                    string headerText = col < tableData[0].Count ? tableData[0][col] : string.Empty;
                    if (string.IsNullOrEmpty(headerText))
                    {
                        headerText = GetColumnName(col);
                    }
                    dataGrid.Columns.Add("col" + col, headerText);
                    dataGrid.Columns[col].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dataGrid.Columns[col].Width = 100;
                }
                dataStartRow = 1;
            }
            else
            {
                for (int col = 0; col < colCount; col++)
                {
                    dataGrid.Columns.Add("col" + col, GetColumnName(col));
                    dataGrid.Columns[col].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dataGrid.Columns[col].Width = 100;
                }
            }

            for (int row = dataStartRow; row < tableData.Count; row++)
            {
                var rowData = new string[colCount];
                for (int col = 0; col < colCount; col++)
                {
                    rowData[col] = col < tableData[row].Count ? tableData[row][col] : string.Empty;
                }
                dataGrid.Rows.Add(rowData);
            }

            dataGrid.RowHeadersWidth = 60;

            UpdateFilterColumnList();
        }

        /// <summary>
        /// フィルタ用の列リストを更新
        /// </summary>
        private void UpdateFilterColumnList()
        {
            cboFilterColumn.Items.Clear();
            cboFilterColumn.Items.Add("(すべての列)");

            for (int col = 0; col < dataGrid.Columns.Count; col++)
            {
                cboFilterColumn.Items.Add(dataGrid.Columns[col].HeaderText);
            }

            if (cboFilterColumn.Items.Count > 0)
            {
                cboFilterColumn.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// 列名を取得（A, B, C, ... Z, AA, AB, ...）
        /// </summary>
        private string GetColumnName(int index)
        {
            string name = string.Empty;
            index++;
            
            while (index > 0)
            {
                index--;
                name = (char)('A' + index % 26) + name;
                index /= 26;
            }
            
            return name;
        }

        /// <summary>
        /// 区切りコンボボックスを更新
        /// </summary>
        private void UpdateDelimiterCombo()
        {
            switch (parser.Delimiter)
            {
                case ',':
                    cboDelimiter.SelectedIndex = 0;
                    break;
                case '\t':
                    cboDelimiter.SelectedIndex = 1;
                    break;
            }
        }

        /// <summary>
        /// 囲みコンボボックスを更新
        /// </summary>
        private void UpdateQuoteCombo()
        {
            cboQuote.SelectedIndex = parser.UseQuotes ? 0 : 1;
        }

        /// <summary>
        /// データ変更イベントを発火
        /// </summary>
        protected virtual void OnDataChanged()
        {
            DataChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// フォントを設定
        /// </summary>
        public void SetFont(Font font)
        {
            dataGrid.DefaultCellStyle.Font = font;
            dataGrid.ColumnHeadersDefaultCellStyle.Font = new Font(font, FontStyle.Bold);
        }

        /// <summary>
        /// 選択列のインデックスを取得
        /// </summary>
        public int GetSelectedColumnIndex()
        {
            if (dataGrid.SelectedCells.Count > 0)
            {
                return dataGrid.SelectedCells[0].ColumnIndex;
            }
            return -1;
        }
    }

    /// <summary>
    /// 列選択イベント引数
    /// </summary>
    public class ColumnSelectedEventArgs : EventArgs
    {
        public int ColumnIndex { get; private set; }
        public string ColumnName { get; private set; }

        public ColumnSelectedEventArgs(int columnIndex, string columnName)
        {
            ColumnIndex = columnIndex;
            ColumnName = columnName;
        }
    }
}
