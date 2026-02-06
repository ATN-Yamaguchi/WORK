using System;
using System.Collections.Generic;
using System.Text;

namespace YamaguchiTextEditor
{
    /// <summary>
    /// CSV/TSVパーサー
    /// </summary>
    public class CsvParser
    {
        /// <summary>
        /// 区切り文字
        /// </summary>
        public char Delimiter { get; set; }

        /// <summary>
        /// 囲み文字を使用するか
        /// </summary>
        public bool UseQuotes { get; set; }

        /// <summary>
        /// 囲み文字
        /// </summary>
        public char QuoteChar { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CsvParser()
        {
            Delimiter = ',';
            UseQuotes = true;
            QuoteChar = '"';
        }

        /// <summary>
        /// テキストをパースして2次元リストに変換
        /// </summary>
        public List<List<string>> Parse(string text)
        {
            var result = new List<List<string>>();
            
            if (string.IsNullOrEmpty(text))
            {
                return result;
            }

            // 改行を統一
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");
            
            var lines = text.Split(new[] { '\n' }, StringSplitOptions.None);
            
            foreach (var line in lines)
            {
                // 最後の空行はスキップ（ただし1行のみの場合は除く）
                if (string.IsNullOrEmpty(line) && result.Count > 0 && lines[lines.Length - 1] == line)
                {
                    continue;
                }
                
                var row = ParseLine(line);
                result.Add(row);
            }

            return result;
        }

        /// <summary>
        /// 1行をパースしてフィールドリストに変換
        /// </summary>
        private List<string> ParseLine(string line)
        {
            var fields = new List<string>();
            
            if (string.IsNullOrEmpty(line))
            {
                fields.Add(string.Empty);
                return fields;
            }

            if (!UseQuotes)
            {
                // 囲み文字なしの場合は単純分割
                var parts = line.Split(new[] { Delimiter });
                fields.AddRange(parts);
                return fields;
            }

            // 囲み文字ありの場合のパース
            var sb = new StringBuilder();
            bool inQuotes = false;
            int i = 0;

            while (i < line.Length)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == QuoteChar)
                    {
                        // 次の文字もクォートならエスケープ
                        if (i + 1 < line.Length && line[i + 1] == QuoteChar)
                        {
                            sb.Append(QuoteChar);
                            i += 2;
                            continue;
                        }
                        else
                        {
                            // クォート終了
                            inQuotes = false;
                            i++;
                            continue;
                        }
                    }
                    else
                    {
                        sb.Append(c);
                        i++;
                    }
                }
                else
                {
                    if (c == QuoteChar)
                    {
                        // クォート開始
                        inQuotes = true;
                        i++;
                    }
                    else if (c == Delimiter)
                    {
                        // フィールド区切り
                        fields.Add(sb.ToString());
                        sb.Clear();
                        i++;
                    }
                    else
                    {
                        sb.Append(c);
                        i++;
                    }
                }
            }

            // 最後のフィールドを追加
            fields.Add(sb.ToString());

            return fields;
        }

        /// <summary>
        /// 2次元リストをテキストに変換
        /// </summary>
        public string ToText(List<List<string>> data)
        {
            if (data == null || data.Count == 0)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();

            for (int row = 0; row < data.Count; row++)
            {
                var rowData = data[row];
                
                for (int col = 0; col < rowData.Count; col++)
                {
                    if (col > 0)
                    {
                        sb.Append(Delimiter);
                    }

                    string field = rowData[col] ?? string.Empty;
                    
                    if (UseQuotes && NeedsQuoting(field))
                    {
                        sb.Append(QuoteChar);
                        sb.Append(field.Replace(QuoteChar.ToString(), new string(QuoteChar, 2)));
                        sb.Append(QuoteChar);
                    }
                    else
                    {
                        sb.Append(field);
                    }
                }

                if (row < data.Count - 1)
                {
                    sb.Append("\n");
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// フィールドがクォートを必要とするか判定
        /// </summary>
        private bool NeedsQuoting(string field)
        {
            if (string.IsNullOrEmpty(field))
            {
                return false;
            }

            return field.Contains(Delimiter.ToString()) ||
                   field.Contains(QuoteChar.ToString()) ||
                   field.Contains("\n") ||
                   field.Contains("\r");
        }

        /// <summary>
        /// 最大列数を取得
        /// </summary>
        public static int GetMaxColumnCount(List<List<string>> data)
        {
            int maxCols = 0;
            foreach (var row in data)
            {
                if (row.Count > maxCols)
                {
                    maxCols = row.Count;
                }
            }
            return maxCols;
        }

        /// <summary>
        /// 列数を揃える（不足分は空文字で埋める）
        /// </summary>
        public static void NormalizeColumnCount(List<List<string>> data)
        {
            int maxCols = GetMaxColumnCount(data);
            
            foreach (var row in data)
            {
                while (row.Count < maxCols)
                {
                    row.Add(string.Empty);
                }
            }
        }
    }
}
