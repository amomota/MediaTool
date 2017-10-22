using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Reflection;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Tar;
using System.Diagnostics;
using System.Data;
using System.Text.RegularExpressions;

namespace Media_Tool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTable m_dt;
        string gzfilePath = "";
        string media_label = "";

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                InitTables();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Title, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // メンバ関数
        /// <summary>
        /// テーブルの初期化
        /// </summary>
        private void InitTables()
        {
            //
            listBox1.Items.Add("解凍するファイルをドラッグアンドドロップしてください");

            //水平スクロールバー
            dataGrid1.HorizontalScrollBarVisibility = ScrollBarVisibility.Visible;

            //垂直スクロールバー
            dataGrid1.VerticalScrollBarVisibility = ScrollBarVisibility.Visible;

            //キー入力を受け付けないようにする
            dataGrid1.IsReadOnly = true;

            m_dt = new DataTable("DataGridTest");

            m_dt.Columns.Add(new DataColumn("file_name", typeof(string)));
            m_dt.Columns.Add(new DataColumn("folder_name", typeof(string)));
            m_dt.Columns.Add(new DataColumn("size", typeof(int)));
        }

        private void listBox1_Drop(object sender, DragEventArgs e)
        {

            // ファイルが渡されていなければ、何もしない
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            // ファイルが渡される度に画面を再描画する。
            m_dt.Rows.Clear();

            // 渡されたファイルに対して処理を行う
            foreach (string filePath in (string[])e.Data.GetData(DataFormats.FileDrop))
            {
                listBox1.Items.Add(filePath);

                string dirPath = System.IO.Path.GetDirectoryName(filePath);
                string stExtension_gz = System.IO.Path.GetExtension(filePath);

                // Mediaのラベルを取得(複数ファイルをドラッグ＆ドロップした時の識別子に使用
                Regex r = new Regex("\\w+_", RegexOptions.IgnoreCase);
                Match m = r.Match(filePath);
                if (m.Success)
                {
                    media_label = m.Value;
                }

                // 抽出したxmlファイルを格納するフォルダを作成
                string xmldirPath = filePath + "_" + System.DateTime.Now.ToString("yyyyMMdd_hhmm");
                System.IO.Directory.CreateDirectory(xmldirPath);

                if (stExtension_gz == ".gz")
                {
                    // gzファイル名退避
                    gzfilePath = filePath;
                    // 解凍インターフェース
                    extract7zip(filePath, dirPath);
                }

                // ファイル名に「Hoge」を含み、拡張子が「.txt」のファイルを最下層まで検索し取得する
                string[] stFilePathes = GetFilesMostDeep(dirPath, "*.cpio", media_label);
                stFilePathes = GetFilesMostDeep(dirPath, "*.rpm", media_label);
                stFilePathes = GetFilesMostDeep(dirPath, "*.cpio", media_label);
                stFilePathes = GetFilesMostDeep(dirPath, "*.xml", media_label);

                DataRow newRowItem;

                // 取得したファイル名を列挙する
                foreach (string stFilePath in stFilePathes)
                {
                    string fn = System.IO.Path.GetFileName(stFilePath);
                    string dn = System.IO.Path.GetDirectoryName(stFilePath);
                    System.IO.FileInfo fi = new System.IO.FileInfo(stFilePath);
                    long filesize = fi.Length;

                    newRowItem = m_dt.NewRow();
                    newRowItem["file_name"] = fn;   //ファイル名 取得
                    newRowItem["folder_name"] = dn; //フォルダ名 取得
                    newRowItem["size"] = filesize;  //ファイルサイズ 取得
                    m_dt.Rows.Add(newRowItem);

                    if ( filesize == 0 )
                    {
                        MessageBox.Show("ファイルサイズ0のファイルがあります。\r\n" + fn);
                    }

                    // 抽出したxmlファイルを準備したフォルダにコピー
                    try
                    {
                        System.IO.File.Copy(stFilePath, xmldirPath + "\\" + fn);
                    }
                    catch(IOException)
                    {
                        MessageBox.Show("同じファイルを２回コピーしていませんか？");
                        this.Activate();
                        MessageBox.Show("処理を終了します");
                        this.Activate();
                        Environment.Exit(0);
                    }
                }
                // グリッドにバインド
                dataGrid1.DataContext = m_dt;

                // 全解凍したフォルダは用済みなので削除
                System.IO.Directory.Delete(gzfilePath.Substring(0, gzfilePath.Length - 8), true);

            }

            // 全解凍したフォルダは用済みなので削除
            //            System.IO.Directory.Delete(gzfilePath.Substring(0, gzfilePath.Length - 8), true);
        }


        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグドロップ時にカーソルの形状を変更
            e.Effects = DragDropEffects.All;
        }

        private static void extract7zip(string fPath,string dPath)
        {

            //7z.exe起動
            ProcessStartInfo startInfo = new ProcessStartInfo(@"C:\Program Files\7-Zip\7z.exe");
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "x -y -o" + @"""" + dPath + "\\" + @"""" + " " + @"""" + fPath + @"""";

            Process proc = Process.Start(startInfo);
            //7zipが終了するまで待ち合わせ
            proc.WaitForExit();
            
        }


        /// ---------------------------------------------------------------------------------------
        /// <summary>
        ///     指定した検索パターンに一致するファイルを最下層まで検索しすべて返します。</summary>
        /// <param name="stRootPath">
        ///     検索を開始する最上層のディレクトリへのパス。</param>
        /// <param name="stPattern">
        ///     パス内のファイル名と対応させる検索文字列。</param>
        /// <returns>
        ///     検索パターンに一致したすべてのファイルパス。</returns>
        /// ---------------------------------------------------------------------------------------
        public static string[] GetFilesMostDeep(string stRootPath, string stPattern, string label)
        {
            System.Collections.Specialized.StringCollection hStringCollection = (
                new System.Collections.Specialized.StringCollection()
            );

            // このディレクトリ内のすべてのファイルを検索する
            foreach (string stFilePath in System.IO.Directory.GetFiles(stRootPath, stPattern))
            {
                // ファイルパスに"config","docs","oem"がある場合は検索せずにスキップする
                if (Regex.IsMatch(stFilePath,"config") || Regex.IsMatch(stFilePath, "docs") || Regex.IsMatch(stFilePath, "oem"))
                {
                    continue;
                }

                // ファイルパスが".cpio.gz_20170719_1220"で終わるフォルダは検索せずにスキップする
                if (Regex.IsMatch(stFilePath, ".cpio.gz_(\\d+)_(\\d+)"))
                {
                    continue;
                }

                if (!Regex.IsMatch(stFilePath,label))
                {
                    continue;
                }

                string stExtension = System.IO.Path.GetExtension(stFilePath);
                string dirPath = System.IO.Path.GetDirectoryName(stFilePath);

                if (stExtension == ".gz")
                {
                    //解凍インターフェース
                    extract7zip(stFilePath, dirPath);
                }
                else if(stExtension == ".cpio" || stExtension == ".rpm")
                {
                    //解凍インターフェース
                    extract7zip(stFilePath, dirPath);
                    System.IO.File.Delete(stFilePath);
                }
                hStringCollection.Add(stFilePath);
            }

            // このディレクトリ内のすべてのサブディレクトリを検索する (再帰)
            foreach (string stDirPath in System.IO.Directory.GetDirectories(stRootPath))
            {
                string[] stFilePathes = GetFilesMostDeep(stDirPath, stPattern, label);

                // 条件に合致したファイルがあった場合は、ArrayList に加える
                if (stFilePathes != null)
                {
                    hStringCollection.AddRange(stFilePathes);
                }
            }

            // StringCollection を 1 次元の String 配列にして返す
            string[] stReturns = new string[hStringCollection.Count];
            hStringCollection.CopyTo(stReturns, 0);
            return stReturns;
        }
    }
}
