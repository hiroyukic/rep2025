using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Access.Dao;
using Application = System.Windows.Forms.Application;

////////////////////////////////////////////////////////////

namespace aeGrep
{
    public partial class Form1 : Form
    {
        //ログオブジェクト
        private Logger log;
        private String iniFileName;

        public Form1()
        {
            InitializeComponent();
            //iniファイルのcomboBox1への設定
            SetComboCBox();
            //iniファイルの読込（初期表示用）
            IniRead();
        }

        //<summary>GrepボタンのAction. Excelファイルを読込みVBAソースを抽出する。<summary>
        private void button1_Click(object sender, EventArgs e)
        {
            //テキストクリア
            richTextBox2.Text = String.Empty;

            //検索条件（ファイル名、検索キー）をiniファイルへ書込み
            IniWrite(textBox1.Text, richTextBox1.Text);


            // 直近のフォルダ名→ログのファイル名称に使用
            string dirName = System.IO.Path.GetFileName(textBox1.Text);

            // ログオブジェクト生成
            log = Logger.GetInstance(dirName);
            log.OutLine(1, "<検索フォルダ>" + textBox1.Text);

            // CSVログタイトル
            log.OutLine(3, "<Dir>, <FileName>, <SearchWord>, <LineNum>, <Line>");
            String outDir = "";
            try
            {
                // 指定ファイルの存在チェック
                if (!System.IO.File.Exists(textBox1.Text))
                {
                    richTextBox2.Text = "【WARNING】指定のファイルが見つかりません。" + Environment.NewLine;
                    return;
                }

                richTextBox2.Text = textBox1.Text + Environment.NewLine;

                // Access/Excel チェック{1:Access, 2:Excel, 9:Other}
                switch (accessExcelCheck(textBox1.Text))
                {



                    case 1:
                        VbaAccess access = new VbaAccess(textBox1.Text);

                        // VBAソースファイルの出力先フォルダの作成
                        outDir = access.GetExportPath(textBox1.Text);

                        /////////////////////////

                        // VBAソースファイルを抽出
                        access.VbaExport4Access(outDir);

                        // クエリをエクスポート
                        access.QueryExport4Access(outDir);

                        access.Dispose();
                        access = null;

                        break;
                    case 2:
                        VbaExcel excel = new VbaExcel();

                        // VBAプロジェクトプロテクションチェック/トラストセンター
                        if (excel.VbaProjectProtectionCheck(textBox1.Text) != true)
                        {
                            richTextBox2.Text = "【WARNING】Excelプロジェクトにパスワードが設定されています。解除魚再実行して下さい。" + Environment.NewLine;
                            return;
                        }
                        // Excelのソースをエクスポート
                        outDir = excel.ExcelVbaExport(textBox1.Text);
                        excel = null;
                        break;
                    default:
                        richTextBox2.Text = "【WARNING】Access/Excel 以外のファイルが設定されています。" + Environment.NewLine;
                        return;
                 }
            }
            catch (Exception ex)
            {
                ///this.Dispose();

                richTextBox2.Text = "【WARNING】" + Environment.NewLine + ex.Message + Environment.NewLine;
                return;
            }
            // エクスポートフォルダパスの表示
            SerchKeyWord(outDir, richTextBox1.Text);

            // ログオブジェクト　Dispose
            log.Close();
            log = null;

            }

        //<summary>Fileボタン　 Excelファイルを指定し画面にセットする<summary>
        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "ファイル選択";
                openFileDialog.Filter = "Excel(*.xlsm:*.xls) | *.xlsm;*xls| Access(*.accdb;*.mdb)|*.accdb;*.mdb)|すべてのファイル(*.*)|*.*";

                //ファイル選択ダイアログを開く
                if(openFileDialog.ShowDialog () == DialogResult.OK)
                {
                    textBox1.Text = openFileDialog.FileName;
                }

            }
        }

        //<summary>テンポラリに存在するiniファイルをリストアップし、コンボボックスに設定する</summary>
        private void SetComboBox()
        {
            string currentDir = Directory.GetCurrentDirectory();
            // iniファイル取り出し
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(currentDir);
            System.IO.FileInfo[] files = di.GetFiles("*.ini", System.IO.SearchOption.TopDirectoryOnly);

            //コンボボックス初期化
            comboBox1.Items.Clear();

            //コンボボックス　アイテム追加
            comboBox1.Items.Add("grep.ini");
            foreach (FileInfo file in files)
            {
                if (file.Name == "grep.ini") continue;

                comboBox1.Items.Add(file);
            }

            // 初期値
            comboBox1.SelectedIndex = 0;
        }


        //<summary>コンボボックスからiniファイルが選択された時のアクション　</summary>
        //<param name = "sender"></param>
        //<param name = "e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            iniFileName = comboBox1.Text;

            //テキストクリア
            textBox1.Text = String.Empty;
            richTextBox2.Text = String.Empty;
            //iniファイル読込み→再描画
            IniRead();
        }


        //<summary>Access/Excelの判定関数<summary>
        //<param name = "filePath">判定対象ファイルのフルパス</param>
        private static int accessExcelCheck(string filePath)
        {
            int ACCESS = 1;
            int EXCEL = 2;
            int OTHER = 9;

            switch (Path.GetExtension (filePath))
            {
                case ".xlsm":
                case ".xlsx":
                case ".xls":
                    return EXCEL;
                case ".accdb":
                case ".mdb":
                    return ACCESS;
            default:
                return OTHER;
            }
        }




        //<summary>iniファイル（検索条件）を書込む<summary>
        //<param name = "key1">Excelファイルのフルパス</param>
        //<param name = "key2">検索キーワード</param>
        private void SetComboCBox()
        {
            String currentDir = Directory.GetCurrentDirectory();
            // C:\\test 以下の、.txtファイルを取得する。（テンポラリのみ）
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(currentDir);
            System.IO.FileInfo[] files = di.GetFiles("*.ini", System.IO.SearchOption.TopDirectoryOnly);

            //コンボボックス初期化
            comboBox1.Items.Clear();

            //コンボボックス　アイテム追加
            comboBox1.Items.Add("grep.ini");
            foreach (FileInfo file in files)
            {
                if (file.Name == "grep.ini") continue;
                comboBox1.Items.Add(file);
            }

            // 初期値
            comboBox1.SelectedIndex = 0;
        }




        //<summary>iniファイル（検索条件）を読込み、画面項目にセットする。<summary>
        private void IniRead()
        {
            // iniファイルのパスを作成（実行ファイルと同じ場所にgrep.ini
            string iniPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + iniFileName;

            // iniファイルのオブジェクト生成
            IniFile ini = new IniFile(iniPath);

            // iniファイルからの読出し
            ini.Section = "Section1";

            string path = ini.GetValue("Key1", "DefaultValue");     //デフォルト値を指定
            string keyWord = ini.GetValue("Key2", "DefaultValue");     //デフォルト値を指定
            bool boolValue = ini.GetValue("Key3", false);           //デフォルト値にfalseを指定
            DateTime datetime = ini.GetValue("Key4", DateTime.MinValue);           //デフォルト値に最小値を指定

            // 検索キーサイズチェック
            if(keyWord.Length > 4000)
            {
                richTextBox2.Text = "検索キーのサイズが4000を超過しています。削除して下さい。" + Environment.NewLine;
                return;
            }

            // 読み出した値を表示
            Console.WriteLine($"Key1: [path]");
            Console.WriteLine($"Key2: [keyWord]");

            // 検索対象フォルダのパス
            textBox1.Text = path;
            // 検索キーワード
            string str1 = keyWord.Replace("<CR>", Environment.NewLine);
            richTextBox1.Text = str1;
        }//IniRead

        //<summary>iniファイル（検索条件）を書込む<summary>
        //<param name = "key1">Excelファイルのフルパス</param>
        //<param name = "key2">検索キーワード</param>
        private void IniWrite(string key1, string key2)
        {
            //iniファイルのパスを作成（実行ファイルを同じ場所にgrep.ini）
            string iniPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + iniFileName;
            //iniファイルのオブジェクト生成
            IniFile ini = new IniFile(iniPath);

            // iniファイルへの書込み
            ini.Section = "Section1";
            string strKey1 = key1;
            ini.SetValue("Key1", strKey1);

            // 改行コードを"<CR>"に変換
            string Keys = key2;
            string strKey2 = Keys.Replace("\r", "").Replace("\n", "<CR>");

            ini.SetValue("Key1", strKey1);
            ini.SetValue("Key2", strKey2);
            ini.SetValue("Key3", true);
            ini.SetValue("Key4", DateTime.Now);
        }//IniWrite


        //<summary>iniファイル（検索条件）を書込む<summary>
        //<param name = "excelPath">Excelファイルのフルパス</param>
        //<param name = "keyWords">検索キーワード</param>
        private void SerchKeyWord(string excelPath, string keyWords)
        {
            //検索キーをリスト化
            string[] wordList = System.Text.RegularExpressions.Regex.Split(keyWords, $"\r\n|\n|\r");

            // Startメッセージ
            richTextBox2.AppendText("- START -" + Environment.NewLine);

            System.IO.DirectoryInfo di = null;
            System.IO.FileInfo[] files = null;

            // 検索キーリストで走査
            foreach (string searchString in wordList)
            {
                // 検索キーがブランクの場合、飛ばす
                if (string.IsNullOrWhiteSpace(searchString)) continue;
                // 先頭文字がシングルコーテーションの場合、コメント行として読み飛ばす。
                if (searchString.StartsWith ("'")) continue;

                richTextBox2.AppendText(searchString + Environment.NewLine);

                try
                {
                    // "C:\test"以下の".txt" ファイルを全て取得する。
                    di = new System.IO.DirectoryInfo(excelPath);
                    files = di.GetFiles("*", System.IO.SearchOption.AllDirectories);

                    string dirPath = "";

                    // 走査（ファイル名）
                    foreach (FileInfo file in files)
                    {
                        // 以下の拡張子は処理しない
                        if (Path.GetExtension(file.Name) == "frx") continue;
                        if (Path.GetExtension(file.Name) == "frm") continue;
                        if (Path.GetExtension(file.Name) == "zip") continue;
                        if (Path.GetExtension(file.Name) == "dll") continue;
                        if (Path.GetExtension(file.Name) == "pdb") continue;
                        if (Path.GetExtension(file.Name) == "xls") continue;
                        if (Path.GetExtension(file.Name) == "xlsx") continue;
                        if (Path.GetExtension(file.Name) == "xlsm") continue;
                        if (Path.GetExtension(file.Name) == "jpg") continue;

                        // 保持しているディレクトリと異なればログに出力
                        if(dirPath != System.IO.Path.GetDirectoryName (file.FullName ))
                        {
                            // 再設定
                            dirPath = System.IO.Path.GetDirectoryName(file.FullName);
                            // ログ出力（ディレクトリパス）
                            Console.WriteLine("");
                            Console.WriteLine(dirPath + "---------------------------");
                        }

                        // --------------------------------------
                        // 検索
                        // --------------------------------------
                        Grep(file.FullName, searchString);

                    }

                }
                catch (Exception ex) {
                    Console.WriteLine("エラーが発生しました。" + ex.Message);
                    richTextBox2.AppendText(Environment.NewLine + "【Error】:" + ex.Message + Environment.NewLine);
                    log = null;
                }
                finally{
                    di = null;
                    files = null;
                }
            }
            // 終了メッセージ
            richTextBox2.AppendText("-END- " + Environment.NewLine);
            richTextBox2.AppendText(Environment.NewLine  + "検索結果：" + Environment.NewLine + log.logFilePath);
        }// void SerchKeyWord(


        //<summary>VBAソースファイルからキーワードを検索する。</summary>
        //<param name = "fullFileName"></param>
        //<param name = "searchWord"></param>
        private bool Grep(string fullFileName, string searchWord)
        {
            string filePath = fullFileName;
            string searchString = searchWord;
            string fileName = System.IO.Path.GetFileName(filePath) ;
            string dirName = System.IO.Path.GetDirectoryName(filePath) ;

            StreamReader sr = null;
            try
            {
                sr = new System.IO.StreamReader(filePath, System.Text.Encoding.GetEncoding("shift_jis"));
                {
                    int nn = 0;
                    string line;
                    while((line = sr.ReadLine()) != null)
                    {
                        nn++;
                        // 大文字小文字区別しないでマッチング
                        if(line.IndexOf(searchString, StringComparison.OrdinalIgnoreCase ) >=0)
                        {
                            //マッチした場合、ログ出力
                            Console.WriteLine(" " + searchWord + "-" + nn + " " + line);

                            line = line.Trim();
                            if(checkBox1.Checked )
                            {
                                //checkBoxがチェックされている場合、カンマをリプレース
                                line = line.Replace(",", "▲");
                            }
                            log.OutLine(3, "," + dirName + "," + fileName + "," + searchWord + "," + nn + "," + line);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("エラーが発生しました" + ex.Message);
                return false;
            }
            finally
            {
                sr.Close();
                sr = null;
            }

            return true;

        } // bool Grep




        }//class Form1 
    }//aeGrep
