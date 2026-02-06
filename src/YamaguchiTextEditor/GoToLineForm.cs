using System;
using System.Drawing;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 指定行へ移動するダイアログ
    /// </summary>
    public class GoToLineForm : Form
    {
        private Label lblLineNumber;
        private TextBox txtLineNumber;
        private Button btnGo;
        private Button btnCancel;

        private int _maxLineNumber;

        /// <summary>
        /// 入力された行番号
        /// </summary>
        public int SelectedLineNumber { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="maxLineNumber">最大行数</param>
        /// <param name="currentLineNumber">現在の行番号</param>
        public GoToLineForm(int maxLineNumber, int currentLineNumber)
        {
            _maxLineNumber = maxLineNumber;
            if (_maxLineNumber < 1)
            {
                _maxLineNumber = 1;
            }

            InitializeComponents();

            txtLineNumber.Text = currentLineNumber.ToString();
            txtLineNumber.SelectAll();
        }

        /// <summary>
        /// コンポーネント初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Text = "指定行へ移動";
            this.Size = new Size(300, 130);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;

            lblLineNumber = new Label();
            lblLineNumber.Text = string.Format("行番号 (1 ～ {0}):", _maxLineNumber);
            lblLineNumber.Location = new Point(12, 15);
            lblLineNumber.Size = new Size(260, 20);
            this.Controls.Add(lblLineNumber);

            txtLineNumber = new TextBox();
            txtLineNumber.Location = new Point(12, 38);
            txtLineNumber.Size = new Size(260, 25);
            txtLineNumber.KeyPress += TxtLineNumber_KeyPress;
            this.Controls.Add(txtLineNumber);

            btnGo = new Button();
            btnGo.Text = "移動";
            btnGo.Location = new Point(106, 70);
            btnGo.Size = new Size(80, 28);
            btnGo.Click += BtnGo_Click;
            this.Controls.Add(btnGo);

            btnCancel = new Button();
            btnCancel.Text = "キャンセル";
            btnCancel.Location = new Point(192, 70);
            btnCancel.Size = new Size(80, 28);
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnGo;
            this.CancelButton = btnCancel;
        }

        /// <summary>
        /// 数字以外の入力を抑制
        /// </summary>
        private void TxtLineNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 数字、バックスペース、コントロールキーのみ許可
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 移動ボタンクリック
        /// </summary>
        private void BtnGo_Click(object sender, EventArgs e)
        {
            string input = txtLineNumber.Text.Trim();
            
            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("行番号を入力してください。", "入力エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtLineNumber.Focus();
                return;
            }

            int lineNumber;
            if (!int.TryParse(input, out lineNumber))
            {
                MessageBox.Show("有効な数値を入力してください。", "入力エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtLineNumber.SelectAll();
                txtLineNumber.Focus();
                return;
            }

            if (lineNumber < 1 || lineNumber > _maxLineNumber)
            {
                MessageBox.Show(
                    string.Format("行番号は 1 から {0} の範囲で入力してください。", _maxLineNumber), 
                    "入力エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtLineNumber.SelectAll();
                txtLineNumber.Focus();
                return;
            }

            SelectedLineNumber = lineNumber;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
