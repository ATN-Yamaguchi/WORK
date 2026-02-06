using System;
using System.Drawing;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 特殊文字（空白、改行）を可視化するRichTextBox
    /// パフォーマンス最適化版
    /// </summary>
    public class SpecialCharRichTextBox : RichTextBox
    {
        // 特殊文字表示設定
        private bool _showSpecialChars = true;
        private Color _specialCharColor = Color.FromArgb(100, 180, 220);

        // 特殊文字の記号
        private const string LF_SYMBOL = "↓";
        private const string CR_SYMBOL = "←";
        private const string FULLWIDTH_SPACE_SYMBOL = "□";
        private const string HALFWIDTH_SPACE_SYMBOL = "_";
        private const string TAB_SYMBOL = "→";

        // 描画用キャッシュ
        private Font _symbolFont;
        private Brush _symbolBrush;
        
        // 描画制御
        private bool _isDrawing = false;
        private Timer _drawTimer;
        private bool _needsRedraw = false;

        /// <summary>
        /// 特殊文字を表示するかどうか
        /// </summary>
        public bool ShowSpecialChars
        {
            get { return _showSpecialChars; }
            set
            {
                if (_showSpecialChars != value)
                {
                    _showSpecialChars = value;
                    RequestRedraw();
                }
            }
        }

        /// <summary>
        /// 特殊文字の表示色
        /// </summary>
        public Color SpecialCharColor
        {
            get { return _specialCharColor; }
            set
            {
                _specialCharColor = value;
                DisposeDrawingResources();
                RequestRedraw();
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SpecialCharRichTextBox()
        {
            // タイマーで描画を遅延実行（パフォーマンス向上）
            _drawTimer = new Timer();
            _drawTimer.Interval = 50;  // 50ms後に描画
            _drawTimer.Tick += DrawTimer_Tick;
        }

        /// <summary>
        /// タイマーによる描画実行
        /// </summary>
        private void DrawTimer_Tick(object sender, EventArgs e)
        {
            _drawTimer.Stop();
            if (_needsRedraw && _showSpecialChars)
            {
                _needsRedraw = false;
                DrawSpecialCharsInternal();
            }
        }

        /// <summary>
        /// 再描画をリクエスト
        /// </summary>
        private void RequestRedraw()
        {
            _needsRedraw = true;
            _drawTimer.Stop();
            _drawTimer.Start();
        }

        /// <summary>
        /// 描画リソースを破棄
        /// </summary>
        private void DisposeDrawingResources()
        {
            if (_symbolFont != null)
            {
                _symbolFont.Dispose();
                _symbolFont = null;
            }
            if (_symbolBrush != null)
            {
                _symbolBrush.Dispose();
                _symbolBrush = null;
            }
        }

        /// <summary>
        /// 描画リソースを取得または作成
        /// </summary>
        private void EnsureDrawingResources()
        {
            if (_symbolFont == null || _symbolFont.Size != this.Font.Size * 0.85f)
            {
                _symbolFont?.Dispose();
                _symbolFont = new Font(this.Font.FontFamily, this.Font.Size * 0.85f, FontStyle.Regular);
            }
            if (_symbolBrush == null)
            {
                _symbolBrush = new SolidBrush(_specialCharColor);
            }
        }

        /// <summary>
        /// 特殊文字を描画（内部実装）
        /// </summary>
        private void DrawSpecialCharsInternal()
        {
            if (_isDrawing || !_showSpecialChars || string.IsNullOrEmpty(this.Text))
            {
                return;
            }

            _isDrawing = true;

            try
            {
                using (Graphics g = this.CreateGraphics())
                {
                    EnsureDrawingResources();

                    string text = this.Text;
                    int textLength = text.Length;
                    
                    // 表示範囲を取得
                    int firstVisibleChar = this.GetCharIndexFromPosition(new Point(0, 0));
                    int lastVisibleChar = this.GetCharIndexFromPosition(new Point(this.ClientSize.Width, this.ClientSize.Height));
                    
                    // 安全マージンを追加
                    int startIndex = Math.Max(0, firstVisibleChar - 50);
                    int endIndex = Math.Min(textLength - 1, lastVisibleChar + 50);

                    // 最大描画文字数を制限（パフォーマンス対策）
                    int maxChars = 2000;
                    if (endIndex - startIndex > maxChars)
                    {
                        endIndex = startIndex + maxChars;
                    }

                    for (int i = startIndex; i <= endIndex && i < textLength; i++)
                    {
                        char c = text[i];
                        string symbol = GetSymbolForChar(c);

                        if (symbol != null)
                        {
                            Point pos = this.GetPositionFromCharIndex(i);

                            // 表示範囲内かチェック
                            if (pos.Y >= -20 && pos.Y <= this.ClientSize.Height + 20 &&
                                pos.X >= -50 && pos.X <= this.ClientSize.Width + 50)
                            {
                                g.DrawString(symbol, _symbolFont, _symbolBrush, pos.X, pos.Y);
                            }
                        }
                    }
                }
            }
            catch
            {
                // 描画エラーは無視
            }
            finally
            {
                _isDrawing = false;
            }
        }

        /// <summary>
        /// 文字に対応する記号を取得
        /// </summary>
        private string GetSymbolForChar(char c)
        {
            switch (c)
            {
                case '\n': return LF_SYMBOL;
                case '\r': return CR_SYMBOL;
                case '\u3000': return FULLWIDTH_SPACE_SYMBOL;
                case ' ': return HALFWIDTH_SPACE_SYMBOL;
                case '\t': return TAB_SYMBOL;
                default: return null;
            }
        }

        /// <summary>
        /// WM_PAINTメッセージ処理
        /// </summary>
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // WM_PAINT後に特殊文字を描画
            if (m.Msg == 0x000F && _showSpecialChars)  // WM_PAINT
            {
                DrawSpecialCharsInternal();
            }
        }

        /// <summary>
        /// スクロール時
        /// </summary>
        protected override void OnVScroll(EventArgs e)
        {
            base.OnVScroll(e);
            if (_showSpecialChars)
            {
                RequestRedraw();
            }
        }

        /// <summary>
        /// 水平スクロール時
        /// </summary>
        protected override void OnHScroll(EventArgs e)
        {
            base.OnHScroll(e);
            if (_showSpecialChars)
            {
                RequestRedraw();
            }
        }

        /// <summary>
        /// リサイズ時
        /// </summary>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_showSpecialChars)
            {
                RequestRedraw();
            }
        }

        /// <summary>
        /// フォント変更時
        /// </summary>
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            DisposeDrawingResources();
            if (_showSpecialChars)
            {
                RequestRedraw();
            }
        }

        /// <summary>
        /// 破棄
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _drawTimer?.Stop();
                _drawTimer?.Dispose();
                DisposeDrawingResources();
            }
            base.Dispose(disposing);
        }
    }
}
