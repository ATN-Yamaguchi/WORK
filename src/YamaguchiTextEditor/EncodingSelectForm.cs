using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 文字コード選択ダイアログ
    /// </summary>
    public class EncodingSelectForm : Form
    {
        private ComboBox cboEncoding;
        private Button btnOK;
        private Button btnCancel;
        private Label lblDescription;

        /// <summary>
        /// 選択された文字コード
        /// </summary>
        public Encoding SelectedEncoding { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public EncodingSelectForm(Encoding currentEncoding = null)
        {
            InitializeComponents();
            
            // 現在の文字コードを選択状態にする
            if (currentEncoding != null)
            {
                SelectEncodingInCombo(currentEncoding);
            }
            else
            {
                cboEncoding.SelectedIndex = 0; // UTF-8
            }
        }

        /// <summary>
        /// コンポーネント初期化
        /// </summary>
        private void InitializeComponents()
        {
            this.Text = "文字コードの選択";
            this.Size = new Size(320, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;

            lblDescription = new Label();
            lblDescription.Text = "文字コードを選択してください:";
            lblDescription.Location = new Point(12, 15);
            lblDescription.Size = new Size(280, 20);
            this.Controls.Add(lblDescription);

            cboEncoding = new ComboBox();
            cboEncoding.DropDownStyle = ComboBoxStyle.DropDownList;
            cboEncoding.Location = new Point(12, 40);
            cboEncoding.Size = new Size(280, 25);
            
            // 文字コードリスト
            cboEncoding.Items.Add(new EncodingItem("UTF-8", new UTF8Encoding(false)));
            cboEncoding.Items.Add(new EncodingItem("UTF-8 (BOM付き)", new UTF8Encoding(true)));
            cboEncoding.Items.Add(new EncodingItem("Shift-JIS", Encoding.GetEncoding(932)));
            cboEncoding.Items.Add(new EncodingItem("EUC-JP", Encoding.GetEncoding(51932)));
            cboEncoding.Items.Add(new EncodingItem("JIS (ISO-2022-JP)", Encoding.GetEncoding(50220)));
            cboEncoding.Items.Add(new EncodingItem("Unicode (UTF-16 LE)", Encoding.Unicode));
            cboEncoding.Items.Add(new EncodingItem("Unicode (UTF-16 BE)", Encoding.BigEndianUnicode));
            
            cboEncoding.SelectedIndex = 0;
            this.Controls.Add(cboEncoding);

            btnOK = new Button();
            btnOK.Text = "OK";
            btnOK.Location = new Point(116, 75);
            btnOK.Size = new Size(85, 28);
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnCancel = new Button();
            btnCancel.Text = "キャンセル";
            btnCancel.Location = new Point(207, 75);
            btnCancel.Size = new Size(85, 28);
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        /// <summary>
        /// 指定された文字コードをコンボボックスで選択
        /// </summary>
        private void SelectEncodingInCombo(Encoding encoding)
        {
            int codePage = encoding.CodePage;
            
            for (int i = 0; i < cboEncoding.Items.Count; i++)
            {
                var item = cboEncoding.Items[i] as EncodingItem;
                if (item != null && item.Encoding.CodePage == codePage)
                {
                    // UTF-8のBOM有無を判定
                    if (codePage == 65001)
                    {
                        var utf8 = encoding as UTF8Encoding;
                        var itemUtf8 = item.Encoding as UTF8Encoding;
                        if (utf8 != null && itemUtf8 != null)
                        {
                            bool hasBom = utf8.GetPreamble().Length > 0;
                            bool itemHasBom = itemUtf8.GetPreamble().Length > 0;
                            if (hasBom == itemHasBom)
                            {
                                cboEncoding.SelectedIndex = i;
                                return;
                            }
                        }
                    }
                    else
                    {
                        cboEncoding.SelectedIndex = i;
                        return;
                    }
                }
            }
            
            cboEncoding.SelectedIndex = 0;
        }

        /// <summary>
        /// OKボタンクリック
        /// </summary>
        private void BtnOK_Click(object sender, EventArgs e)
        {
            var selected = cboEncoding.SelectedItem as EncodingItem;
            if (selected != null)
            {
                SelectedEncoding = selected.Encoding;
            }
            else
            {
                SelectedEncoding = new UTF8Encoding(false);
            }
            
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

        /// <summary>
        /// 文字コードアイテム（コンボボックス用）
        /// </summary>
        private class EncodingItem
        {
            public string DisplayName { get; private set; }
            public Encoding Encoding { get; private set; }

            public EncodingItem(string displayName, Encoding encoding)
            {
                DisplayName = displayName;
                Encoding = encoding;
            }

            public override string ToString()
            {
                return DisplayName;
            }
        }
    }
}
