using System;
using System.Windows.Forms;

namespace YamaguchiTextEditor
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメインエントリーポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // コマンドライン引数でファイルパスが渡された場合
            string initialFilePath = null;
            if (args != null && args.Length > 0 && !string.IsNullOrEmpty(args[0]))
            {
                initialFilePath = args[0];
            }

            Application.Run(new MainForm(initialFilePath));
        }
    }
}
