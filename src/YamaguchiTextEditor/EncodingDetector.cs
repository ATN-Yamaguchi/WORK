using System;
using System.IO;
using System.Text;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// 文字コード自動判定クラス
    /// </summary>
    public class EncodingDetector
    {
        /// <summary>
        /// 判定結果
        /// </summary>
        public class DetectionResult
        {
            /// <summary>
            /// 検出された文字コード
            /// </summary>
            public Encoding Encoding { get; set; }

            /// <summary>
            /// 判定に成功したか
            /// </summary>
            public bool IsDetected { get; set; }

            /// <summary>
            /// 判定方法の説明
            /// </summary>
            public string DetectionMethod { get; set; }

            public DetectionResult()
            {
                Encoding = new UTF8Encoding(false);
                IsDetected = false;
                DetectionMethod = "デフォルト";
            }
        }

        /// <summary>
        /// ファイルの文字コードを自動判定
        /// </summary>
        public static DetectionResult DetectEncoding(string filePath)
        {
            var result = new DetectionResult();

            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);
                return DetectEncoding(bytes);
            }
            catch (Exception)
            {
                return result;
            }
        }

        /// <summary>
        /// バイト配列から文字コードを自動判定
        /// </summary>
        public static DetectionResult DetectEncoding(byte[] bytes)
        {
            var result = new DetectionResult();

            if (bytes == null || bytes.Length == 0)
            {
                return result;
            }

            // 1. BOMチェック
            var bomResult = DetectByBOM(bytes);
            if (bomResult != null)
            {
                result.Encoding = bomResult;
                result.IsDetected = true;
                result.DetectionMethod = "BOM検出";
                return result;
            }

            // 2. JISチェック（ESCシーケンス）
            if (IsJIS(bytes))
            {
                result.Encoding = Encoding.GetEncoding(50220);
                result.IsDetected = true;
                result.DetectionMethod = "JIS(ISO-2022-JP)パターン検出";
                return result;
            }

            // 3. UTF-8チェック（BOMなし）
            if (IsValidUtf8(bytes))
            {
                // ASCII範囲外の文字が含まれているかチェック
                bool hasNonAscii = false;
                foreach (byte b in bytes)
                {
                    if (b >= 0x80)
                    {
                        hasNonAscii = true;
                        break;
                    }
                }

                if (hasNonAscii)
                {
                    result.Encoding = new UTF8Encoding(false);
                    result.IsDetected = true;
                    result.DetectionMethod = "UTF-8パターン検出";
                    return result;
                }
            }

            // 4. EUC-JPチェック
            int eucScore = GetEucJpScore(bytes);
            
            // 5. Shift-JISチェック
            int sjisScore = GetShiftJisScore(bytes);

            // スコアを比較
            if (eucScore > 0 || sjisScore > 0)
            {
                if (eucScore > sjisScore)
                {
                    result.Encoding = Encoding.GetEncoding(51932);
                    result.IsDetected = true;
                    result.DetectionMethod = "EUC-JPパターン検出";
                    return result;
                }
                else if (sjisScore > 0)
                {
                    result.Encoding = Encoding.GetEncoding(932);
                    result.IsDetected = true;
                    result.DetectionMethod = "Shift-JISパターン検出";
                    return result;
                }
            }

            // 6. 純粋なASCIIの場合
            bool isPureAscii = true;
            foreach (byte b in bytes)
            {
                if (b >= 0x80)
                {
                    isPureAscii = false;
                    break;
                }
            }

            if (isPureAscii)
            {
                result.Encoding = new UTF8Encoding(false);
                result.IsDetected = true;
                result.DetectionMethod = "ASCII（UTF-8互換）";
                return result;
            }

            // 判定できない場合
            result.Encoding = new UTF8Encoding(false);
            result.IsDetected = false;
            result.DetectionMethod = "判定不能";
            return result;
        }

        /// <summary>
        /// BOMから文字コードを検出
        /// </summary>
        private static Encoding DetectByBOM(byte[] bytes)
        {
            if (bytes.Length >= 3)
            {
                // UTF-8 BOM: EF BB BF
                if (bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                {
                    return new UTF8Encoding(true);
                }
            }

            if (bytes.Length >= 2)
            {
                // UTF-16 LE BOM: FF FE
                if (bytes[0] == 0xFF && bytes[1] == 0xFE)
                {
                    return Encoding.Unicode;
                }

                // UTF-16 BE BOM: FE FF
                if (bytes[0] == 0xFE && bytes[1] == 0xFF)
                {
                    return Encoding.BigEndianUnicode;
                }
            }

            return null;
        }

        /// <summary>
        /// JIS（ISO-2022-JP）かどうかをチェック
        /// </summary>
        private static bool IsJIS(byte[] bytes)
        {
            // JISはESCシーケンスを使用する
            // ESC $ B (1B 24 42) - JIS X 0208
            // ESC $ @ (1B 24 40) - JIS X 0208 (旧)
            // ESC ( B (1B 28 42) - ASCII
            // ESC ( J (1B 28 4A) - JIS X 0201 Roman

            for (int i = 0; i < bytes.Length - 2; i++)
            {
                if (bytes[i] == 0x1B) // ESC
                {
                    if (i + 2 < bytes.Length)
                    {
                        // ESC $ B or ESC $ @
                        if (bytes[i + 1] == 0x24 && (bytes[i + 2] == 0x42 || bytes[i + 2] == 0x40))
                        {
                            return true;
                        }
                        // ESC ( B or ESC ( J
                        if (bytes[i + 1] == 0x28 && (bytes[i + 2] == 0x42 || bytes[i + 2] == 0x4A))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 有効なUTF-8かどうかをチェック
        /// </summary>
        private static bool IsValidUtf8(byte[] bytes)
        {
            int i = 0;
            while (i < bytes.Length)
            {
                byte b = bytes[i];

                int continuationBytes = 0;

                if (b <= 0x7F)
                {
                    // ASCII (0xxxxxxx)
                    i++;
                    continue;
                }
                else if ((b & 0xE0) == 0xC0)
                {
                    // 2バイト文字 (110xxxxx)
                    continuationBytes = 1;
                }
                else if ((b & 0xF0) == 0xE0)
                {
                    // 3バイト文字 (1110xxxx)
                    continuationBytes = 2;
                }
                else if ((b & 0xF8) == 0xF0)
                {
                    // 4バイト文字 (11110xxx)
                    continuationBytes = 3;
                }
                else
                {
                    // 不正なUTF-8
                    return false;
                }

                // 継続バイトをチェック
                for (int j = 0; j < continuationBytes; j++)
                {
                    i++;
                    if (i >= bytes.Length)
                    {
                        return false;
                    }
                    if ((bytes[i] & 0xC0) != 0x80)
                    {
                        return false;
                    }
                }

                i++;
            }

            return true;
        }

        /// <summary>
        /// EUC-JPらしさのスコアを計算
        /// </summary>
        private static int GetEucJpScore(byte[] bytes)
        {
            int score = 0;
            int i = 0;

            while (i < bytes.Length)
            {
                byte b = bytes[i];

                // EUC-JP: 2バイト文字は 0xA1-0xFE の連続
                if (b >= 0xA1 && b <= 0xFE)
                {
                    if (i + 1 < bytes.Length)
                    {
                        byte b2 = bytes[i + 1];
                        if (b2 >= 0xA1 && b2 <= 0xFE)
                        {
                            score += 2;
                            i += 2;
                            continue;
                        }
                    }
                }

                // EUC-JP: 半角カナは 0x8E + 0xA1-0xDF
                if (b == 0x8E)
                {
                    if (i + 1 < bytes.Length)
                    {
                        byte b2 = bytes[i + 1];
                        if (b2 >= 0xA1 && b2 <= 0xDF)
                        {
                            score += 1;
                            i += 2;
                            continue;
                        }
                    }
                }

                // 不正なバイトがあればスコアを下げる
                if (b >= 0x80 && b <= 0xA0)
                {
                    score -= 1;
                }

                i++;
            }

            return score;
        }

        /// <summary>
        /// Shift-JISらしさのスコアを計算
        /// </summary>
        private static int GetShiftJisScore(byte[] bytes)
        {
            int score = 0;
            int i = 0;

            while (i < bytes.Length)
            {
                byte b = bytes[i];

                // Shift-JIS: 第1バイト 0x81-0x9F, 0xE0-0xFC
                if ((b >= 0x81 && b <= 0x9F) || (b >= 0xE0 && b <= 0xFC))
                {
                    if (i + 1 < bytes.Length)
                    {
                        byte b2 = bytes[i + 1];
                        // 第2バイト 0x40-0x7E, 0x80-0xFC
                        if ((b2 >= 0x40 && b2 <= 0x7E) || (b2 >= 0x80 && b2 <= 0xFC))
                        {
                            score += 2;
                            i += 2;
                            continue;
                        }
                    }
                }

                // 半角カナ 0xA1-0xDF
                if (b >= 0xA1 && b <= 0xDF)
                {
                    score += 1;
                    i++;
                    continue;
                }

                // 不正なバイトがあればスコアを下げる
                if (b >= 0xFD)
                {
                    score -= 1;
                }

                i++;
            }

            return score;
        }

        /// <summary>
        /// 文字コードの表示名を取得
        /// </summary>
        public static string GetEncodingDisplayName(Encoding encoding)
        {
            if (encoding == null)
            {
                return "不明";
            }

            int codePage = encoding.CodePage;

            if (codePage == 65001)
            {
                var utf8 = encoding as UTF8Encoding;
                if (utf8 != null && utf8.GetPreamble().Length > 0)
                {
                    return "UTF-8 (BOM付き)";
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
                    return "JIS (ISO-2022-JP)";
                case 1200:
                    return "UTF-16 LE";
                case 1201:
                    return "UTF-16 BE";
                default:
                    return encoding.EncodingName;
            }
        }
    }
}
