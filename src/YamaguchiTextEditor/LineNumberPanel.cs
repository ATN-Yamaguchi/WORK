using System;
using System.Drawing;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 行番号を表示するカスタムパネル
    /// </summary>
    public class LineNumberPanel : Panel
    {
        private RichTextBox _targetTextBox;
        private Font _font;
        private Color _foreColor;
        private Color _backColor;
        private int _padding;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public LineNumberPanel()
        {
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | 
                          ControlStyles.AllPaintingInWmPaint | 
                          ControlStyles.UserPaint, true);
            
            _font = new Font("MS ゴシック", 10.0f, FontStyle.Regular);
            _foreColor = Color.DimGray;
            _backColor = Color.FromArgb(245, 245, 245);
            _padding = 5;
            
            this.BackColor = _backColor;
        }

        /// <summary>
        /// 対象のRichTextBoxを設定
        /// </summary>
        public void SetTargetTextBox(RichTextBox textBox)
        {
            _targetTextBox = textBox;
        }

        /// <summary>
        /// フォントを設定
        /// </summary>
        public void SetFont(Font font)
        {
            if (font != null)
            {
                _font = new Font(font.FontFamily, font.Size, FontStyle.Regular);
                this.Invalidate();
            }
        }

        /// <summary>
        /// 必要な幅を計算
        /// </summary>
        public int CalculateRequiredWidth()
        {
            if (_targetTextBox == null)
            {
                return 50;
            }

            int lineCount = _targetTextBox.Lines.Length;
            if (lineCount < 1)
            {
                lineCount = 1;
            }

            string maxLineNumber = lineCount.ToString();
            
            using (Graphics g = this.CreateGraphics())
            {
                SizeF size = g.MeasureString(maxLineNumber, _font);
                return (int)size.Width + (_padding * 2) + 10;
            }
        }

        /// <summary>
        /// 再描画
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (_targetTextBox == null)
            {
                return;
            }

            Graphics g = e.Graphics;
            g.Clear(_backColor);

            // RichTextBoxの最初の表示行を取得
            int firstCharIndex = _targetTextBox.GetCharIndexFromPosition(new Point(0, 0));
            int firstLine = _targetTextBox.GetLineFromCharIndex(firstCharIndex);

            // 1行の高さを取得
            float lineHeight = _font.GetHeight(g);
            
            // RichTextBoxからの実際の行高さを使用
            int testCharIndex = _targetTextBox.GetFirstCharIndexFromLine(firstLine);
            Point testPoint = Point.Empty;
            if (testCharIndex >= 0)
            {
                testPoint = _targetTextBox.GetPositionFromCharIndex(testCharIndex);
            }

            int totalLines = _targetTextBox.Lines.Length;
            if (totalLines < 1)
            {
                totalLines = 1;
            }

            using (Brush brush = new SolidBrush(_foreColor))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Far;
                sf.LineAlignment = StringAlignment.Center;

                int y = testPoint.Y;
                int lineNumber = firstLine + 1;

                while (y < this.Height && lineNumber <= totalLines)
                {
                    // 次の行の開始位置を取得して行の高さを計算
                    int currentLineCharIndex = _targetTextBox.GetFirstCharIndexFromLine(lineNumber - 1);
                    int nextLineCharIndex = -1;
                    
                    if (lineNumber < totalLines)
                    {
                        nextLineCharIndex = _targetTextBox.GetFirstCharIndexFromLine(lineNumber);
                    }

                    int currentLineHeight;
                    if (nextLineCharIndex > 0)
                    {
                        Point currentPos = _targetTextBox.GetPositionFromCharIndex(currentLineCharIndex);
                        Point nextPos = _targetTextBox.GetPositionFromCharIndex(nextLineCharIndex);
                        currentLineHeight = nextPos.Y - currentPos.Y;
                        if (currentLineHeight < lineHeight)
                        {
                            currentLineHeight = (int)lineHeight;
                        }
                    }
                    else
                    {
                        currentLineHeight = (int)lineHeight;
                    }

                    // 行番号を描画
                    RectangleF rect = new RectangleF(
                        0, 
                        y, 
                        this.Width - _padding, 
                        currentLineHeight
                    );
                    
                    g.DrawString(lineNumber.ToString(), _font, brush, rect, sf);

                    y += currentLineHeight;
                    lineNumber++;
                }
            }

            // 右端に区切り線を描画
            using (Pen pen = new Pen(Color.LightGray, 1))
            {
                g.DrawLine(pen, this.Width - 1, 0, this.Width - 1, this.Height);
            }
        }

        /// <summary>
        /// リソース解放
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_font != null)
                {
                    _font.Dispose();
                    _font = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
