using System;
using System.Diagnostics;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // 実行ファイルのベース名を取得
        string exePath = Process.GetCurrentProcess().MainModule.FileName;
        string baseName = Path.GetFileNameWithoutExtension(exePath);
        string cmdFileName = baseName + ".cmd";

        // ProcessStartInfoを作成
        ProcessStartInfo psi = new ProcessStartInfo();
        psi.FileName = cmdFileName;
        psi.CreateNoWindow = true; // コマンドプロンプトのウィンドウを表示しない
        psi.UseShellExecute = false;

        // コマンドを実行する
        try
        {
            Process p = Process.Start(psi);
            p.WaitForExit();
        }
        catch (Exception ex)
        {
            Console.WriteLine("エラーが発生しました: " + ex.Message);
        }
    }
}