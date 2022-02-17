using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FrontEnd
{
    class GeneralCSV
    {
        private Dictionary<string, string> CsvToDict;

        public GeneralCSV(string filePath)
        {

            CsvToDict = new Dictionary<string, string>();
            XSSFWorkbook xssfwb; // initialize excel reader


            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // Opening spreadsheet with spreadsheet reader 
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet("Conditions");

            for (int row = 0; row < sheet.PhysicalNumberOfRows; row++)
            {
                string key = sheet.GetRow(row).GetCell(0).ToString();
                string value;

                if (sheet.GetRow(row).GetCell(1) != null)
                {
                    value = sheet.GetRow(row).GetCell(1).ToString();
                }
                else
                {
                    value = " ";
                }

                CsvToDict.Add(key, value);
            }

            xssfwb.Close();

        }

        public bool DictContains(string key)
        {
            return CsvToDict.ContainsKey(key);
        }

        public string getValue(string key)
        {
            return CsvToDict[key];

        }

        public string getKey(string key)
        {
            return key;
        }
    }
}


