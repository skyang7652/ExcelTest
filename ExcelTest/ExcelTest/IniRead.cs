using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
    class IniRead
    {
        public string FileName; //INI文件名
        //聲明讀寫INI文件的API函數
        [DllImport("kernel32")]
        private static extern bool WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, byte[] retVal, int size, string filePath);

        //類的構造函數，傳遞INI文件名
        public IniRead(string AFileName)
        {
            // 判斷文件是否存在
            FileInfo fileInfo = new FileInfo(AFileName);
            //Todo:搞清枚舉的用法
            if ((!fileInfo.Exists))
            { //|| (FileAttributes.Directory in fileInfo.Attributes))
                //文件不存在，建立文件
                System.IO.StreamWriter sw = new System.IO.StreamWriter(AFileName, false, System.Text.Encoding.Default);
                try
                {
                    sw.Write("#表格配置檔案");
                    sw.Close();
                }
                catch
                {
                    throw (new ApplicationException("Ini文件不存在"));
                }
            }
            //必須是完全路徑，不能是相對路徑
            FileName = fileInfo.FullName;
        }

        //寫INI文件
        public void WriteString(string Section, string Ident, string Value)
        {
            if (!WritePrivateProfileString(Section, Ident, Value, FileName))
            {

                throw (new ApplicationException("寫Ini文件出錯"));
            }
        }

        //讀取INI文件指定
        public string ReadString(string Section, string Ident, string Default)
        {
            Byte[] Buffer = new Byte[65535];
            int bufLen = GetPrivateProfileString(Section, Ident, Default, Buffer, Buffer.GetUpperBound(0), FileName);
            //必須設定0（系統預設的代碼頁）的編碼方式，否則無法支持中文
            string s = Encoding.GetEncoding(0).GetString(Buffer);
            s = s.Substring(0, bufLen);
            return s.Trim();
        }

        //讀整數
        public int ReadInteger(string Section, string Ident, int Default)
        {
            string intStr = ReadString(Section, Ident, Convert.ToString(Default));
            try
            {
                return Convert.ToInt32(intStr);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Default;
            }
        }

        //寫整數
        public void WriteInteger(string Section, string Ident, int Value)
        {
            WriteString(Section, Ident, Value.ToString());
        }

        //讀布爾
        public bool ReadBool(string Section, string Ident, bool Default)
        {
            try
            {
                return Convert.ToBoolean(ReadString(Section, Ident, Convert.ToString(Default)));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Default;
            }
        }

        //寫Bool
        public void WriteBool(string Section, string Ident, bool Value)
        {
            WriteString(Section, Ident, Convert.ToString(Value));
        }

        //從Ini文件中，將指定的Section名稱中的所有Ident添加到列表中
        public void ReadSection(string Section, StringCollection Idents)
        {
            Byte[] Buffer = new Byte[16384];
            //Idents.Clear();

            int bufLen = GetPrivateProfileString(Section, null, null, Buffer, Buffer.GetUpperBound(0), FileName);
            //對Section進行解析
            GetStringsFromBuffer(Buffer, bufLen, Idents);
        }

        private void GetStringsFromBuffer(Byte[] Buffer, int bufLen, StringCollection Strings)
        {
            Strings.Clear();
            if (bufLen != 0)
            {
                int start = 0;
                for (int i = 0; i < bufLen; i++)
                {
                    if ((Buffer[i] == 0) && ((i - start) > 0))
                    {
                        String s = Encoding.GetEncoding(0).GetString(Buffer, start, i - start);
                        Strings.Add(s);
                        start = i + 1;
                    }
                }
            }
        }

        //從Ini文件中，讀取所有的Sections的名稱
        public void ReadSections(StringCollection SectionList)
        {
            //Note:必須得用Bytes來實現，StringBuilder只能取到第一個Section
            byte[] Buffer = new byte[65535];
            int bufLen = 0;
            bufLen = GetPrivateProfileString(null, null, null, Buffer,
            Buffer.GetUpperBound(0), FileName);
            GetStringsFromBuffer(Buffer, bufLen, SectionList);
        }

        //讀取指定的Section的所有Value到列表中
        public void ReadSectionValues(string Section, NameValueCollection Values)
        {
            StringCollection KeyList = new StringCollection();
            ReadSection(Section, KeyList);
            Values.Clear();
            foreach (string key in KeyList)
            {
                Values.Add(key, ReadString(Section, key, ""));
            }
        }

        /**/
        ////讀取指定的Section的所有Value到列表中，
        /*public void ReadSectionValues(string Section, NameValueCollection Values,char splitString)
        {　 string sectionValue;
        　　string[] sectionValueSplit;
        　　StringCollection KeyList = new StringCollection();
        　　ReadSection(Section, KeyList);
        　　Values.Clear();
        　　foreach (string key in KeyList)
        　　{
        　　　　sectionValue=ReadString(Section, key, "");
        　　　　sectionValueSplit=sectionValue.Split(splitString);
        　　　　Values.Add(key, sectionValueSplit[0].ToString(),sectionValueSplit[1].ToString());
        　　}
        }*/

        //清除某個Section
        public void EraseSection(string Section)
        {
            if (!WritePrivateProfileString(Section, null, null, FileName))
            {
                throw (new ApplicationException("無法清除Ini文件中的Section"));
            }
        }

        //刪除某個Section下的鍵
        public void DeleteKey(string Section, string Ident)
        {
            WritePrivateProfileString(Section, Ident, null, FileName);
        }

        //Note:對於Win9X，來說需要實現UpdateFile方法將緩衝中的數據寫入文件
        //在Win NT, 2000和XP上，都是直接寫文件，沒有緩衝，所以，無須實現UpdateFile
        //執行完對Ini文件的修改之後，應該調用本方法更新緩衝區。
        public void UpdateFile()
        {
            WritePrivateProfileString(null, null, null, FileName);
        }

        //檢查某個Section下的某個鍵值是否存在
        public bool ValueExists(string Section, string Ident)
        {
            StringCollection Idents = new StringCollection();
            ReadSection(Section, Idents);
            return Idents.IndexOf(Ident) > -1;
        }

        //確保資源的釋放
        ~IniRead()
        {
            UpdateFile();
        }






    }
}
