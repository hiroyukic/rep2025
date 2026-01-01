using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace aeGrep
{
    class IniFile
    {
        // iniファイルのパス
        public string FilePath { get; private set; }

        // 処理対象のセクション
        public string Section { get; set; }

        // コンストラクタ
        public IniFile(string path)
        {
            FilePath = path;
            Section = null;
        }

        // 値のセット
        public void SetValue<T>(string key, T Value)
        {
            if(Section == null)
            {
                throw new InvalidOperationException("セクションが指定されていません");
            }

            // サポートしている型あんら処理
            if (IsSupportedType(typeof(T)))
            {
                string strValue = Value.ToString();
                WritePrivateProfileString(Section, key, strValue, FilePath);
            }
            else
            {
                throw new InvalidOperationException($"未対応の型です。：[typeof(T)]");
            }
        }

        // 値の取得
        public T GetValue<T>(string key , T defaultValue)
        {
            if(Section == null)
            {
                throw new InvalidOperationException("セクションが指定されていません");
            }

            // 文字列として取得
            var sb = new StringBuilder(496);
            GetPrivateProfileString(Section, key, "", sb, sb.Capacity, FilePath);
            string value = sb.ToString().Trim();

            // 取得失敗時、デフォルト値
            if(string.IsNullOrEmpty(value))
            {
                return defaultValue;
            }

            // サポートしている型なら処理
            if (IsSupportedType(typeof(T)))
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            else
            {
                throw new InvalidOperationException($"未対応の型です。：");
            }

        }

        // サポートしている型か？
        bool IsSupportedType (Type targetType)
        {
            return targetType.IsPrimitive ||
            targetType == typeof(string) ||
            targetType == typeof(decimal) ||
            targetType == typeof(DateTime);
        }

        //private bool IsSupportdType(Type type)
        //{
        //    throw new NotImplementedException();
        //}

        //private void GetPrivateProfileString(string section, string key, string v, StringBuilder sb, int capacity, string filePath)
        //{
        //    throw new NotImplementedException();
        //}






        // API
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName); 

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        //[Serializable]
        //private class InvalidOperationExcetion : Exception
        //{
        //    public InvalidOperationExcetion()
        //    {
        //    }

        //    public InvalidOperationExcetion(string message) : base(message)
        //    {
        //    }

        //    public InvalidOperationExcetion(string message, Exception innerException) : base(message, innerException)
        //    {
        //    }

        //    protected InvalidOperationExcetion(SerializationInfo info, StreamingContext context) : base(info, context)
        //    {
        //    }
        //}
    }

}
