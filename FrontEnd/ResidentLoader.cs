using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FrontEnd
{
    class ResidentLoader
    {

        private List<string> resList;

        public ResidentLoader(string filePath, int sheetIndex)
        {
            resList = new List<string>();

           
            XSSFWorkbook xssfwb; // initialize excel reader


            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // Opening spreadsheet with spreadsheet reader 
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheetAt(sheetIndex);

            for (int row = 0; row < sheet.PhysicalNumberOfRows; row++)
            {
                resList.Add(sheet.GetRow(row).GetCell(0).ToString());
            }

            xssfwb.Close();


        }

        public List<string> returnList()
        {
            return resList;
        }

            
    }
}
