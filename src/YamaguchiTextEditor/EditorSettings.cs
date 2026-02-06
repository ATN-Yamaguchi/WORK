using System;
using System.Drawing;
using System.Text;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// エディタの設定を保持するクラス
    /// </summary>
    public class EditorSettings
    {
        /// <summary>
        /// エディタのフォント
        /// </summary>
        public Font EditorFont { get; set; }

        /// <summary>
        /// 現在の文字コード
        /// </summary>
        public Encoding CurrentEncoding { get; set; }

        /// <summary>
        /// 現在編集中のファイルパス（新規の場合はnull）
        /// </summary>
        public string CurrentFilePath { get; set; }

        /// <summary>
        /// テキストが変更されたかどうか
        /// </summary>
        public bool IsModified { get; set; }

        /// <summary>
        /// 検出された改行コード
        /// </summary>
        public NewLineType DetectedNewLine { get; set; }

        /// <summary>
        /// コンストラクタ（デフォルト値設定）
        /// </summary>
        public EditorSettings()
        {
            EditorFont = new Font("MS ゴシック", 10.0f, FontStyle.Regular);
            CurrentEncoding = new UTF8Encoding(false); // UTF-8 BOMなし
            CurrentFilePath = null;
            IsModified = false;
            DetectedNewLine = NewLineType.CRLF;
        }

        /// <summary>
        /// 改行コードの表示名を取得
        /// </summary>
        public string GetNewLineDisplayName()
        {
            switch (DetectedNewLine)
            {
                case NewLineType.CRLF:
                    return "CRLF";
                case NewLineType.LF:
                    return "LF";
                case NewLineType.CR:
                    return "CR";
                default:
                    return "CRLF";
            }
        }

        /// <summary>
        /// 改行コード文字列を取得
        /// </summary>
        public string GetNewLineString()
        {
            switch (DetectedNewLine)
            {
                case NewLineType.CRLF:
                    return "\r\n";
                case NewLineType.LF:
                    return "\n";
                case NewLineType.CR:
                    return "\r";
                default:
                    return "\r\n";
            }
        }

        /// <summary>
        /// 文字コードの表示名を取得
        /// </summary>
        public string GetEncodingDisplayName()
        {
            if (CurrentEncoding == null)
            {
                return "UTF-8";
            }

            int codePage = CurrentEncoding.CodePage;

            // BOM付きUTF-8の判定
            if (codePage == 65001)
            {
                var utf8 = CurrentEncoding as UTF8Encoding;
                if (utf8 != null)
                {
                    byte[] preamble = utf8.GetPreamble();
                    if (preamble != null && preamble.Length > 0)
                    {
                        return "UTF-8 (BOM)";
                    }
                }
                return "UTF-8";
            }

            switch (codePage)
            {
                case 932:
                    return "Shift-JIS";
                case 51932:
                    return "EUC-JP";
                case 50220:
                    return "JIS";
                case 1200:
                    return "UTF-16 LE";
                case 1201:
                    return "UTF-16 BE";
                default:
                    return CurrentEncoding.EncodingName;
            }
        }
    }

    /// <summary>
    /// 改行コードの種類
    /// </summary>
    public enum NewLineType
    {
        CRLF,   // Windows標準
        LF,     // Unix/Linux/Mac
        CR      // 旧Mac
    }
}
