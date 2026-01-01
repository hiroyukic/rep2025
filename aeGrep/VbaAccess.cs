using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

using System.IO;

using System.Diagnostics;
//using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Access.Dao;
//using Application = System.Windows.Forms.Application;


namespace aeGrep
{
    class VbaAccess
    {

        static Microsoft.Office.Interop.Access.Application accessApp; // { get; set; }

        // コンストラクタ
        public VbaAccess(string accessPath)
        {
            //
            accessApp = new Microsoft.Office.Interop.Access.Application();
            accessApp.OpenCurrentDatabase(accessPath, false);
            accessApp.Visible = false;
        }

        // コンストラクタ
       public void Dispose()
        {
            accessApp.Visible = true;   //バックグラウンドプロセスに残留させないための処理
            accessApp.CloseCurrentDatabase();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(accessApp);
            accessApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }


        ////<summary>ExcelファイルについてVBAソースを出力する。</summary>
        ////<param name = "ExcelPath">抽出対象のExcelファイルのフルパス</param>
        ////<returns> Ｅｘｐｏｒｔ先フォルダのパスを返す。</returns>
        ////public string AccessVbaExport(string accessPath)
        //public string AccessVbaExport(string exportDirPath)
        //{
        //    //String exportDirPath = "";
        //    //FileInfo accessFile = null;
        //    try
        //    {
        //        //accessFile = new FileInfo(accessPath);
        //        //string accessName = accessFile.Name;
        //        //string dirPath = accessFile.DirectoryName;

        //        ////出力ファイル用のフォルダ作成
        //        //exportDirPath = dirPath + "\\src_" + accessName;
        //        //Directory.CreateDirectory(exportDirPath);

        //        // VBAソース抽出／出力
        //        VbaExport4Access( exportDirPath);

        //        //出力フォルダのパスを返す
        //        return exportDirPath;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine(ex.Message);
        //        //accessApp.Visible = true;
        //        //accessApp.CloseCurrentDatabase();

        //        throw ex;
        //    }
        //    finally
        //    {
        //        accessFile = null;
        //    }
        //}//ExcelVbaExport




        //<summary>ExcelファイルについてVBAソースを出力する。</summary>
        //<param name = "ExcelPath">抽出対象のExcelファイルのフルパス</param>
        //<returns> Ｅｘｐｏｒｔ先フォルダのパスを返す。</returns>
        public void VbaExport4Access( string outPath)
        {
            //String TARGET = accessFilePath;
            string pathName;

            //var accessApp = new Microsoft.Office.Interop.Access.Application();
            //accessApp.OpenCurrentDatabase (TARGET, false);
            //accessApp.Visible = false;

            try
            {



                foreach (VBComponent vbc in accessApp.VBE.ActiveVBProject.VBComponents)
                {
                    var module = vbc.CodeModule;

                    //------------------------------------------------
                    // vbaファイル出力
                    //------------------------------------------------

                    Microsoft.Vbe.Interop.vbext_ComponentType type;
                    type = module.Parent.Type;
                    pathName = Path.Combine(outPath, module.Name);

                    switch (type)
                    {
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                            pathName += ".cls";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                            pathName += ".frm";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                            pathName += ".cls";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                            pathName += ".bas";
                            break;
                        default:
                            pathName += ".bas";
                            break;
                    }
                    //ファイル出力
                    vbc.Export(pathName);
                }

                ///////////////////////////////////////
            //accessApp.Visible = true;

            //    accessApp.CloseCurrentDatabase();
            //    accessApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                this.Dispose();
                throw ex; //20251228
            }
            finally
            {
            }

        }//VbaExport4Access

        //<summary>ExcelファイルについてVBAソースを出力する。</summary>
        //<param name = "ExcelPath">抽出対象のExcelファイルのフルパス</param>
        //<returns> Ｅｘｐｏｒｔ先フォルダのパスを返す。</returns>
        //public void QueryExport4Access(string accessFilePath, string outPath)
        public void QueryExport4Access( string outPath)
        {
            //var accessApp = new Microsoft.Office.Interop.Access.Application();
            //accessApp.OpenCurrentDatabase(TARGET, false);
            accessApp.Visible = false;
            Microsoft.Office.Interop.Access.Dao.Database currentDb = null;

            //クエリ抽出
            string filePath = Path.Combine(outPath, "QUERY.txt");
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                //Microsoft.Office.Interop.Access.Dao.Database currentDb = null;
                currentDb = accessApp.CurrentDb();

                foreach (QueryDef queryDef in currentDb.QueryDefs)
                {
                    sw.WriteLine(queryDef.Name);
                    sw.WriteLine(queryDef.SQL);
                }
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(currentDb);
                //currentDb = null;
            }
            currentDb.Close();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(currentDb);
            currentDb = null;

            ///////////////////////////////////////
            //accessApp.Visible = true;
            //accessApp.CloseCurrentDatabase();

            //accessApp.Quit();

            //後処理
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(accessApp );
            //accessApp = null;

            //GC.Collect();
            //GC.WaitForPendingFinalizers();

        } //void QueryExport4Access

        //<summary>VBAソースの出力先パスを返す。Accessパスの直下に、"src_" +。Accessファイル名のディレクトリを作成する。</summary>
        //<param name = "ExcelPath">抽出対象のAccessファイルのフルパス</param>
        //<returns>出力先フォルダのパスを返す</returns>
        public string GetExportPath(string accessPath)
        {
            String exportDirPath ;
            FileInfo accessFile ;
            try
            {
                accessFile = new FileInfo(accessPath);
                string accessName = accessFile.Name;
                string dirPath = accessFile.DirectoryName;

                //出力ファイル用のフォルダ作成
                exportDirPath = dirPath + "\\src_" + accessName;
                Directory.CreateDirectory(exportDirPath);

                //出力フォルダのパスを返す
                return exportDirPath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }
        }　//GetExportPath







    }// class VbaAccess
}//aeGrep
