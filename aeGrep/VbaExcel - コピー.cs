using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;


//12:00




using System.IO;

using System.Diagnostics;
//using System.Threading;
//using System.Runtime.InteropServices;
//using Microsoft.Vbe.Interop;
//using Microsoft.Office.Interop.Access.Dao;
//using Application = System.Windows.Forms.Application;

namespace aeGrep
{
    class VbaExcel
    {
        // コンストラクタ
        public VbaExcel ()
        {
            //
        }

        //<summary>ExcelファイルについてVBAソースを出力する。</summary>
        //<param name = "ExcelPath">抽出対象のExcelファイルのフルパス</param>
        //<returns> Ｅｘｐｏｒｔ先フォルダのパスを返す。</returns>
        public string ExcelVbaExport(string excelPath)
        {
            try
            {
                String exportDirPath = "";

                FileInfo excelFile = new FileInfo(excelPath);
                string excelName = excelFile.Name;      //"sample.txt"
                string dirPath = excelFile.DirectoryName;

                //出力ファイル用のフォルダ作成
                exportDirPath = dirPath + "\\src_" + excelName;
                Directory.CreateDirectory(exportDirPath);

                // VBAソース抽出／出力
                VbaExport(excelPath, exportDirPath);
                return exportDirPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                throw ex;
            }
            finally
            {
                //
            }
        }

        //<summary>ExcelファイルについてVBAソースを抽出する</summary>
        //<param name = "ExcelFileName">抽出対象のExcelファイルのフルパス</param>
        //<param name = "OutPath">抽出ファイルの出力フォルダ</param>
        public void VbaExport(string excelFileName, String outPath)
        {
            String TARGET;
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbooks books = null;
            Microsoft.Office.Interop.Excel.Workbook book = null;
            
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.EnableEvents = false;
                books = excel.Workbooks;
                TARGET = excelFileName;
                book = books.Open(TARGET, ReadOnly: true, IgnoreReadOnlyRecommended: true);

                book.Activate();

                //VBAファイル抽出
                Microsoft.Vbe.Interop.VBComponents modules = book.VBProject.VBComponents;
                string pathName;

                foreach (Microsoft.Vbe.Interop.VBComponent module in modules)
                {
                    pathName = Path.Combine(outPath, module.Name);
                    switch (module.Type)
                    {
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                            pathName += ".cls";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                            pathName += ".frm";
                            break;
                        default:
                            pathName += ".bas";
                            break;
                    }
                    module.Export(pathName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally
            {
                excel.DisplayAlerts = false;
                book.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                book = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                books.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(books);
                books = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }//vbaExport

        //<summary>Excel VBAプロジェクトにパスワードが設定されているかチェックを実施する。</summary>
        //<param name = "ExcelPath">チェック対象のExcelファイルのフルパス</param>
        //<returns>true/false</returns>
        public Boolean VbaProjectProtectionCheck(string excelPath)
        {
            Boolean ret = false;
            String TARGET;
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbooks books = null;
            Microsoft.Office.Interop.Excel.Workbook book = null;
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.EnableEvents = false;
                books = excel.Workbooks;

                TARGET = excelPath;
                book = books.Open(TARGET, ReadOnly: true, IgnoreReadOnlyRecommended: true);
                // ロックされているかチェック
                if(book.VBProject.Protection == Microsoft.Vbe.Interop.vbext_ProjectProtection.vbext_pp_locked)
                {
                    ret = false;
                }
                else
                {
                    ret = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excel.DisplayAlerts = false;
                book.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                book = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                books.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(books);
                books = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return ret;

        } //boolean VbaProjectProtectionCheck


    } // class VbaExcel
} // aeGrep
