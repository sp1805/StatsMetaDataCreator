using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatMetaDataHelper
{
    public class Program
    {

        public static Excel.Workbook mWorkBook;
        public static Excel.Sheets mWorkSheets;
        public static Excel.Worksheet mWSheet1;
        public static Excel.Range celLrangE;
        public static Excel.Application oXL = new Excel.Application();
        static void Main(string[] args)
        {
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Add(Type.Missing);

            //mWSheet1 = (Excel.Worksheet)mWorkBook.ActiveSheet;
            //mWSheet1.Name = "Data";

            Console.WriteLine("Paste all your files at C:\\StatXmlHelper and remove state specific names from file: Eg- States.RB_Bangalore.xml to States.xml");
            string fileName;
            List<string> desc = new List<string>();
            List<string> dbValue = new List<string>();
            Console.WriteLine("Enter the file name: \n");
            string file = Console.ReadLine();
            fileName = file + ".xml";
            string filePath = @"C:\StatXmlHelper";
            string workBookFileName = @"C:\StatXmlHelper\" + file +".xlsx";
            string DefaultFilePath = System.IO.Path.Combine(filePath, fileName);
            if (File.Exists(DefaultFilePath))
            {
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();

                try
                {
                    xmlDoc.Load(DefaultFilePath);
                    XmlNodeList xnViolationInfos = xmlDoc.SelectNodes("DropDownItems/DropDownItem");

                    if (xnViolationInfos != null)
                    {
                        foreach (XmlNode xnViolationInfo in xnViolationInfos)
                        {
                            desc.Add(xnViolationInfo.SelectSingleNode("Desc").InnerText);
                            dbValue.Add(xnViolationInfo.SelectSingleNode("Data").InnerText);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                try

                {

                    mWSheet1 = mWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data

                    //Writing data into excel of 100 rows with 10 column 

                    for (int r = 1; r < desc.Count; r++) //r stands for ExcelRow and c for ExcelColumn

                    {
                        mWSheet1.Cells[r, 1] = desc[r];
                        mWSheet1.Cells[r, 2] = dbValue[r];
                        mWSheet1.Cells[r, 3] = "<item desc=\"" + desc[r] + "\" dbValue=\"" + dbValue[r] + "\" />";

                    }
                    mWorkBook.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet

                    mWorkBook.SaveAs(workBookFileName);

                    mWorkBook.Close();

                    oXL.Quit();
                }
                catch (Exception exHandle)

                {

                    Console.WriteLine("Exception: " + exHandle.Message);

                    Console.ReadLine();

                }

                finally

                {



                    foreach (Process process in Process.GetProcessesByName("Excel"))

                        process.Kill();

                }
            }
        }
    }
}
