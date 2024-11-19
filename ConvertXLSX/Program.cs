using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Security.Cryptography;


namespace ConvertXLSX
{
    public class TableLine
    {
        public List<string> _stringList = new List<string>();
    }

    public class EnumData
    {
        public class EnumValueData
        {
            public string _name;
            public string _value;
        }

        public string _name;
        public List<EnumValueData> _enumValueList = new List<EnumValueData>();
    }

    public class Program
    {
        static string _encrpytKey = "12345678";
        static string _crpytoAlgorithm = "DES"; // DES, AES256CBC
        static void Main(string[] args)
        {
            Console.WriteLine("Version : 1.4");

            _encrpytKey = args[0];
            _crpytoAlgorithm = args[1];
            string srcPath = args[2];
            string table = args[3];
            string targetPath = args[4];
            string targetPath2 = (args.Length >= 6) ? args[5] : "none";
            string targetPath3 = (args.Length >= 7) ? args[6] : "none";

            if (args.Length >= 3)
            {
                if (srcPath != "none" && !Directory.Exists(srcPath))
                {
                    Console.WriteLine("srcPath does not exist!");
                    return;
                }
                if (targetPath2 != "none" && !Directory.Exists(targetPath2))
                {
                    Console.WriteLine("targetPath2 does not exist!");
                    return;
                }
                if (targetPath3 != "none" && !Directory.Exists(targetPath3))
                {
                    Console.WriteLine("targetPath3 does not exist!");
                    return;
                }
            }

            Dictionary<int, string> dicWrite = new Dictionary<int, string>();
            dicWrite.Clear();

            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(srcPath);

                string searchPattern = table + ".xlsx";
                FileInfo[] fileInfos = directoryInfo.GetFiles(searchPattern);

                bool splitTable = false;
                if (table.Contains("(*)"))
                {
                    splitTable = true;
                    table = table.Replace("(*)", "");
                }

                for (int i = 0; i < fileInfos.Length; i++) 
                {
                    bool append = (splitTable && i > 0);
                    bool writeBinary = (i == fileInfos.Length - 1);

                    FileInfo fileInfo = fileInfos[i];

                    if (fileInfo.Name.Contains("~$") == false)
                    {
                        dicWrite.Clear();

                        Console.WriteLine("source file : " + fileInfo.FullName);

                        if (targetPath != "none")
                        {
                            string targetName = targetPath + table;
                            ExcelToText(fileInfo.FullName, targetName, OutputFileType.Binary, append, writeBinary);
                            
                            Console.WriteLine("target file : " + targetName);
                        }

                        if (targetPath2 != "none")
                        {
                            string targetName = targetPath2 + "Ref_" + table;
                            ExcelToText(fileInfo.FullName, targetName, OutputFileType.Text, append, false);
                            Console.WriteLine("target2 file : " + targetName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("Done!!");
        }

        private static Excel.Workbook excelWorkBook = null;
        private static Excel.Application excelApp = null;
        private static Excel.Worksheet excelWorkSheet = null;

        public enum OutputFileType
        {
            Binary = 1,
            Text
        }

        private static BindingList<TableLine> ExcelToText(string excelFileName, string targetFileName, OutputFileType outputFileType, bool append, bool writeBinary)
        {
            excelApp = new Excel.Application();
            excelApp.Visible = false;
            List<EnumData> enumDataList = new List<EnumData>();
            BindingList<TableLine> tableLineList = new BindingList<TableLine>();
            int lastColumn, lastRow;

            try
            {
                excelWorkBook = excelApp.Workbooks.Open(excelFileName);

                for (int sheet = 1; sheet <= excelWorkBook.Sheets.Count; sheet++)
                {
                    excelWorkSheet = (Excel.Worksheet)excelWorkBook.Sheets[sheet];
                    if (excelWorkSheet.Name == "_Enum")
                    {
                        Excel.Range usedRange = excelWorkSheet.UsedRange;

                        lastColumn = usedRange.Columns.Count;
                        lastRow = usedRange.Rows.Count;
                        Array values = (Array)excelWorkSheet.UsedRange.Cells.Value;

                        for (int i = 1; i <= lastColumn; i += 2)
                        {
                            EnumData enumData = new EnumData();
                            for (int j = 1; j <= lastRow; j++)
                            {
                                if (values.GetValue(j, i) != null)
                                {
                                    if (j == 1) enumData._name = values.GetValue(j, i).ToString();
                                    else
                                    {
                                        EnumData.EnumValueData enumValueData = new EnumData.EnumValueData();
                                        enumValueData._name = values.GetValue(j, i).ToString();
                                        enumValueData._value = values.GetValue(j, i + 1).ToString();
                                        enumData._enumValueList.Add(enumValueData);
                                    }
                                }
                            }
                            enumDataList.Add(enumData);
                        }
                    }
                }

                for (int sheet = 1; sheet <= excelWorkBook.Sheets.Count; sheet++)
                {
                    excelWorkSheet = (Excel.Worksheet)excelWorkBook.Sheets[sheet];
                    if (excelWorkSheet.Name[0] != '_')
                    {
                        Excel.Range UsedRange = excelWorkSheet.UsedRange;
                        Array values = (Array)excelWorkSheet.UsedRange.Cells.Value;

                        // Find the last row with actual data
                        int actualLastRow = 0;
                        // Scan from bottom to top to find the last row with data
                        for (int j = UsedRange.Rows.Count; j >= 1; j--)
                        {
                            bool hasData = false;
                            for (int i = 1; i <= UsedRange.Columns.Count; i++)
                            {
                                if (values.GetValue(j, i) != null && !string.IsNullOrWhiteSpace(values.GetValue(j, i).ToString()))
                                {
                                    hasData = true;
                                    break;
                                }
                            }

                            if (hasData)
                            {
                                actualLastRow = j;
                                break;
                            }
                        }

                        // Find the last column with actual data
                        int actualLastColumn = 0;
                        // Scan from right to left to find the last column with data
                        for (int i = UsedRange.Columns.Count; i >= 1; i--)
                        {
                            bool hasData = false;
                            for (int j = 1; j <= actualLastRow; j++)
                            {
                                if (values.GetValue(j, i) != null && !string.IsNullOrWhiteSpace(values.GetValue(j, i).ToString()))
                                {
                                    hasData = true;
                                    break;
                                }
                            }

                            if (hasData)
                            {
                                actualLastColumn = i;
                                break;
                            }
                        }

                        tableLineList.Clear();
                        // Use actualLastRow and actualLastColumn
                        for (int j = 1; j <= actualLastRow; j++)
                        {
                            TableLine tableLine = new TableLine();
                            for (int i = 1; i <= actualLastColumn; i++)
                            {
                                if (values.GetValue(j, i) != null)
                                {
                                    string value = values.GetValue(j, i).ToString();
                                    if (value == "")
                                    {
                                        value = "Null";
                                        Console.WriteLine(string.Format("Null Data : Row({0}), Column({1})", j, i));
                                    }
                                    tableLine._stringList.Add(value);
                                }
                                else
                                {
                                    Console.WriteLine(string.Format("Invalid Data : Row({0}), Column({1})", j, i));
                                }
                            }
                            tableLineList.Add(tableLine);
                        }

                        string textFileName;
                        if (outputFileType == OutputFileType.Binary)
                        {
                            textFileName = targetFileName + "_" + excelWorkSheet.Name + ".bytes";
                            WriteToFile(textFileName, excelWorkSheet.Name, tableLineList, enumDataList, outputFileType, append, writeBinary);
                        }
                        else if (outputFileType == OutputFileType.Text)
                        {
                            if (sheet == 1)
                            {
                                textFileName = targetFileName + ".txt";
                                WriteToFile(textFileName, excelWorkSheet.Name, tableLineList, enumDataList, outputFileType, append,  false);
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception : " + e.Message);
                excelWorkBook.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(excelWorkBook);
                excelWorkBook = null;
                GC.Collect();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                GC.Collect();

                return tableLineList;
            }

            excelWorkBook.Saved = true;
            excelWorkBook.Close();
            excelApp.Quit();

            Marshal.ReleaseComObject(excelWorkBook);
            excelWorkBook = null;
            GC.Collect();
            Marshal.ReleaseComObject(excelApp);
            excelApp = null;
            GC.Collect();

            return tableLineList;
        }

        static void WriteToStreamText(StreamWriter streamWriter, BindingList<TableLine> tableLineList, List<EnumData> enumDataList, int startIndex)
        {
            string strLine = enumDataList.Count + "\r\n";
            streamWriter.Write(strLine);
            for (int i = 0; i < enumDataList.Count; i++)
            {
                EnumData enumData = enumDataList[i];

                strLine = enumData._name + "\t" + enumData._enumValueList.Count + "\r\n";
                streamWriter.Write(strLine);

                for (int j = 0; j < enumData._enumValueList.Count; j++)
                {
                    EnumData.EnumValueData enumValueData = enumData._enumValueList[j];
                    strLine = enumValueData._name + "\t" + enumValueData._value + "\r\n";
                    streamWriter.Write(strLine);
                }
            }

            if (startIndex == 0)
            {
                strLine = "";
                for (int k = 0; k < tableLineList[0]._stringList.Count; k++)
                {
                    string str = tableLineList[0]._stringList[k];
                    if (str[0] != '#')
                    {
                        strLine += str;
                        if (k < tableLineList[0]._stringList.Count - 1) strLine += "\t";
                    }
                    if (k == tableLineList[0]._stringList.Count - 1) strLine += "\r\n";
                }

                streamWriter.Write(strLine);
            }


            for (int i = 1; i < tableLineList.Count; i++)
            {
                strLine = "";

                for (int k = 0; k < tableLineList[i]._stringList.Count; k++)
                {
                    string typeStr = tableLineList[0]._stringList[k];
                    if (typeStr[0] != '#')
                    {
                        string str = tableLineList[i]._stringList[k];

                        strLine += str;
                        if (k < tableLineList[i]._stringList.Count - 1) strLine += "\t";
                    }
                    if (k == tableLineList[i]._stringList.Count - 1) strLine += "\r\n";
                }

                streamWriter.Write(strLine);
            }            
        }

        public static void WriteToFile(string fileName, string workSheetName, 
            BindingList<TableLine> tableLineList, List<EnumData> enumDataList, OutputFileType outputFileType, bool append, bool writeBinary)
        {
            if (outputFileType == OutputFileType.Binary)
            {
                StreamWriter streamWriter = new StreamWriter(fileName + ".txt", append);
                WriteToStreamText(streamWriter, tableLineList, enumDataList, 1);
                streamWriter.Flush();
                streamWriter.Close();

                if (writeBinary)
                {
                    if (_crpytoAlgorithm == "DES")
                    {
                        FileStream textFileStream = new FileStream(fileName + ".txt", FileMode.Open, FileAccess.Read, FileShare.None);
                        FileStream encryptedFileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);

                        DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                        des.Key = ASCIIEncoding.ASCII.GetBytes(_encrpytKey);
                        des.IV = ASCIIEncoding.ASCII.GetBytes(_encrpytKey);

                        ICryptoTransform encryptTransform = des.CreateEncryptor();
                        CryptoStream cryptoStream = new CryptoStream(encryptedFileStream, encryptTransform, CryptoStreamMode.Write);

                        byte[] fileContents = new byte[textFileStream.Length];
                        textFileStream.Read(fileContents, 0, fileContents.Length);
                        textFileStream.Close();

                        cryptoStream.Write(fileContents, 0, fileContents.Length);
                        encryptedFileStream.Flush();
                        cryptoStream.Close();

                        File.Delete(fileName + ".txt");
                    }
                    else if (_crpytoAlgorithm == "AES256CBC")
                    {
                        FileStream textFileStream = new FileStream(fileName + ".txt", FileMode.Open, FileAccess.Read, FileShare.None);
                        FileStream encryptedFileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);

                        RijndaelManaged aes = new RijndaelManaged();
                        aes.KeySize = 256;
                        aes.BlockSize = 128;
                        aes.Mode = CipherMode.CBC;
                        aes.Padding = PaddingMode.PKCS7;
                        aes.Key = Encoding.ASCII.GetBytes(_encrpytKey);
                        aes.IV = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

                        ICryptoTransform encryptTransform = aes.CreateEncryptor();
                        CryptoStream cryptoStream = new CryptoStream(encryptedFileStream, encryptTransform, CryptoStreamMode.Write);

                        byte[] fileContents = new byte[textFileStream.Length];
                        textFileStream.Read(fileContents, 0, fileContents.Length);
                        textFileStream.Close();

                        cryptoStream.Write(fileContents, 0, fileContents.Length);
                        encryptedFileStream.Flush();
                        cryptoStream.Close();

                        File.Delete(fileName + ".txt");
                    }
                }                
            }
            else if(outputFileType == OutputFileType.Text)
            {
                StreamWriter streamWriter = new StreamWriter(fileName, append);
                int startIndex = (append) ? 1 : 0;
                WriteToStreamText(streamWriter, tableLineList, enumDataList, startIndex);
                streamWriter.Flush();
                streamWriter.Close();
            }
        }
    }
}
