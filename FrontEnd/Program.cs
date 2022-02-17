using System;
using System.IO;
using System.Collections.Generic;
using System.Data.Odbc;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Geom;
using iText.Layout.Borders;
using Table = iText.Layout.Element.Table;
using System.Text.RegularExpressions;


namespace FrontEnd
{
    class Program
    {


        public Program(string filePath, int schedNum, string location)
        {

            string layout = @"C:\Users\whelanb\Downloads\Layout_Template.xlsx";

            GeneralCSV notesByName = new GeneralCSV(@"C:\Users\whelanb\Documents\ConditionalSheets\NotesByName.xlsx"); // notesByName maps specific notes that are applied on a per person basis 
            GeneralCSV wording = new GeneralCSV(@"C:\Users\whelanb\Documents\ConditionalSheets\AssignmentsWithDifferentWording.xlsx"); // wording maps lightning-bolt assignments to a different wording under assignments in the daysheet 
            GeneralCSV assignToNotes = new GeneralCSV(@"C:\Users\whelanb\Documents\ConditionalSheets\AssignmentsMoveToNotes.xlsx"); // assignToNotes maps assignments in lightning-bolt to the notes tab in the daysheet
            GeneralCSV AdditionalNotes = new GeneralCSV(@"C:\Users\whelanb\Documents\ConditionalSheets\AdditionalNotes.xlsx"); // AdditionalNotes maps lightning-bolt assignments to additional notes that should be added to a rows notes

            List<string> jobs = new List<string>(); // Jobs list used to manage job label sections in the daysheet pdf
            List<string> arrAssign = new List<string>(); // arrAssign list used to contain multiple assignments due to multiple lightning-bolt rows being applied to a single staff member
            List<string> arrNotes = new List<string>(); // arrNotes used to contain multiple notes for a single staff member
            List<SheetLine> sheetList = new List<SheetLine>();

            HSSFWorkbook hssfwb; // initialize excel reader
            XSSFWorkbook xssfwb;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read)) // Opening spreadsheet with spreadsheet reader 
            {
                hssfwb = new HSSFWorkbook(file);
                file.Close();
                file.Dispose();

            }



            ISheet sheet = hssfwb.GetSheet("Report"); // Get Report Sheet from xlsx
            // Get Report Sheet from xlsx


            string destDate = sheet.GetRow(2).GetCell(6).ToString();
            DateTime date = DateTime.Parse(destDate);
            destDate = date.ToString("MMM") + "_" + date.ToString("dd") + "_" + date.ToString("yy");
            string dest = location + "\\DaySheet_" + destDate + ".xlsx";


            File.Copy(layout, @dest);









            SheetLine sheetLine; // Declaring sheetline object to used to manage lightning-bolt schedule row data




            for (int row = 1; row < sheet.PhysicalNumberOfRows - 1; row++)
            {

                sheetLine = new SheetLine(); // Creates new sheetLine object for eacg row in the spreadsheet

                if ((row < sheet.LastRowNum) && sheet.GetRow(row).GetCell(2).ToString() == sheet.GetRow(row + 1).GetCell(2).ToString()) // Checks if the name of the staff member of the current row is also in the next row
                {                                                                                                                       // If true, condition adds the assignment and notes of the current row to the arrAssign
                    arrAssign.Add(sheet.GetRow(row).GetCell(3).ToString());                                                             // and arrNotes Lists respectively, and does not add the line to the pdf
                    arrNotes.Add(sheet.GetRow(row).GetCell(4).ToString());
                }

                else
                {

                    for (int column = 0; column < 6; column++) // Loops the the 5 columns of relevent data in the spreadsheet
                    {




                        if (column == 0)
                        {
                            sheetLine.setPost(sheet.GetRow(row).GetCell(column).ToString()); // assigns the 0th column data to sheetline->post
                        }





                        if (column == 1)
                        {
                            sheetLine.setOnCall(sheet.GetRow(row).GetCell(column).ToString()); // assigns the 1st column data to sheetline->OnCall
                            //sheet.
                        }





                        if (column == 2)
                        {
                            string name = sheet.GetRow(row).GetCell(column).ToString();
                            name = name.Replace(",", "");

                            string[] nameArr = name.Split(' ');
                            string finalName = "";

                            for (int i = 0; i < nameArr.Length; i++)
                            {
                                if (nameArr[i].Length != 0)
                                {
                                    if (IsAllUpper(nameArr[i]))
                                    {
                                        name = nameArr[i].ToLower();
                                        name = (char.ToUpper(name[0])).ToString() + name.Substring(1);
                                        finalName = finalName + name + " ";

                                    }
                                }
                            }
                            finalName = finalName.Trim();

                            if (notesByName.DictContains(finalName))
                            {
                                arrNotes.Add(notesByName.getValue(finalName));
                            }

                            sheetLine.setStaff(sheet.GetRow(row).GetCell(column).ToString());
                        }




                        if (column == 3)
                        {
                            string assign = "";
                            string toNote;
                            for (int i = 0; i < arrAssign.Count; i++)
                            {
                                if (wording.DictContains(arrAssign[i]))
                                {
                                    toNote = wording.getValue(arrAssign[i]);

                                    if (assignToNotes.DictContains(toNote))
                                    {
                                        arrNotes.Add(assignToNotes.getValue(toNote));
                                    }
                                    else
                                    {
                                        assign = assign + toNote + " ";
                                    }
                                }
                                else
                                {
                                    assign = assign + arrAssign[i] + " ";
                                }

                            }

                            if (wording.DictContains(sheet.GetRow(row).GetCell(column).ToString()))
                            {
                                toNote = wording.getValue(sheet.GetRow(row).GetCell(column).ToString());



                                if (assignToNotes.DictContains(toNote))
                                {
                                    string str = assignToNotes.getValue(toNote);
                                    arrNotes.Add(str);

                                }

                                else if (AdditionalNotes.DictContains(sheet.GetRow(row).GetCell(column).ToString()))
                                {
                                    string str = AdditionalNotes.getValue(sheet.GetRow(row).GetCell(column).ToString());

                                    arrNotes.Add(str);
                                    str = wording.getValue(sheet.GetRow(row).GetCell(column).ToString());

                                    assign = assign + str + " ";

                                }

                                else
                                {
                                    assign = assign + toNote + " ";
                                }
                            }

                            else if (assignToNotes.DictContains(sheet.GetRow(row).GetCell(column).ToString()))
                            {
                                toNote = assignToNotes.getValue(sheet.GetRow(row).GetCell(column).ToString());
                                arrNotes.Add(toNote);
                            }




                            else
                            {
                                assign = assign + sheet.GetRow(row).GetCell(column).ToString() + " ";
                            }


                            sheetLine.setAssignment(assign);
                        }





                        if (column == 4)
                        {
                            string attachNote = "";
                            if (arrNotes.Count > 0)
                            {
                                for (int i = 0; i < arrNotes.Count; i++)
                                {
                                    attachNote = attachNote + arrNotes[i] + " ";
                                }
                            }
                            attachNote = attachNote + sheet.GetRow(row).GetCell(column).ToString();
                            sheetLine.setNotes(attachNote);
                        }

                        if (column == 5)
                        {
                            sheetLine.setPType(sheet.GetRow(row).GetCell(column).ToString());
                        }

                    }

                    sheetList.Add(sheetLine);
                    arrAssign.Clear();
                    arrNotes.Clear();
                }


            }

            hssfwb.Close();


            using (FileStream file = new FileStream(dest, FileMode.Open, FileAccess.ReadWrite)) // Opening spreadsheet with spreadsheet reader 
            {
                xssfwb = new XSSFWorkbook(file);

            }

            sheet = xssfwb.GetSheet("Monday");

            

            var style1 = xssfwb.CreateCellStyle();
            style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            

            var style2 = xssfwb.CreateCellStyle();
            style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style2.BorderLeft = BorderStyle.Thin;
            style2.BorderRight = BorderStyle.Thin;
            var font = xssfwb.CreateFont();
            font.FontHeightInPoints = 10;
            style2.SetFont(font);

            var style3 = xssfwb.CreateCellStyle();
            style3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style3.BorderLeft = BorderStyle.Thin;
            style3.BorderRight = BorderStyle.Thin;
            style3.BorderBottom = BorderStyle.Thin;

            var style4 = xssfwb.CreateCellStyle();
            style4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            style4.BorderLeft = BorderStyle.Thin;
            style4.BorderBottom = BorderStyle.Dotted;

            var style5 = xssfwb.CreateCellStyle();
            style5.BorderBottom = BorderStyle.Dotted;

            var dateStyle = xssfwb.CreateCellStyle();
            dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            var dateFont = xssfwb.CreateFont();
            dateFont.IsBold = true;
            dateFont.FontHeightInPoints = 18;

            dateStyle.SetFont(dateFont);


            bool assocFlag = true;
            bool assistFlag = true;
            bool fellowFlag = true;
            bool resFlag = true;
            int writeRow = 4;

            int[] datePos = { 0, 72, 73, 141 };

            for (int i = 0; i < datePos.Length; i++)
            {
                IRow dateRows = sheet.CreateRow(datePos[i]);
                //dateRows.HeightInPoints = 35.25F;
                ICell dayCell = dateRows.CreateCell(2);
                ICell monthCell = dateRows.CreateCell(5);
                ICell numCell = dateRows.CreateCell(6);
                ICell yearCell = dateRows.CreateCell(7);

                dayCell.SetCellValue(date.ToString("ddd"));
                monthCell.SetCellValue(date.ToString("MMM"));
                numCell.SetCellValue(date.ToString("dd"));
                yearCell.SetCellValue(date.ToString("yyyy"));

                dayCell.CellStyle = dateStyle;
                monthCell.CellStyle = dateStyle;
                numCell.CellStyle = dateStyle;
                yearCell.CellStyle = dateStyle;

                dayCell.CellStyle = dateStyle;
                monthCell.CellStyle = dateStyle;
                numCell.CellStyle = dateStyle;
                yearCell.CellStyle = dateStyle;
            }

            
            
            
            
            
            
            for (int i = 0; i < sheetList.Count; i++)
            {

                if ((sheetList[i].getPType() == "Associate") && assocFlag == true)
                {
                    writeRow = 55;
                    assocFlag = false;
                }

                if ((sheetList[i].getPType() == "Anesthesia Assistant") && assistFlag == true)
                {
                    writeRow = 60;
                    assistFlag = false;
                }

                if ((sheetList[i].getPType() == "Fellow") && fellowFlag == true)
                {
                    writeRow = 77;
                    fellowFlag = false;
                }

                if ((sheetList[i].getPType() == "Resident") && resFlag == true)
                {
                    writeRow = 105;
                    resFlag = false;
                }

                IRow dataRow = sheet.CreateRow(writeRow);
                dataRow.HeightInPoints = 14.25F;
                
                ICell cell = dataRow.CreateCell(0);
                ICell cell2 = dataRow.CreateCell(2);
                ICell spacer = dataRow.CreateCell(4);
                ICell cell3 = dataRow.CreateCell(5);
                ICell cell4 = dataRow.CreateCell(6);
                ICell cell5 = dataRow.CreateCell(7);
                ICell noteExtOne = dataRow.CreateCell(8);
                ICell noteExtTwo = dataRow.CreateCell(9);
                
                cell.SetCellValue(sheetList[i].getPost());
                cell2.SetCellValue(sheetList[i].getOnCall());
                cell3.SetCellValue(sheetList[i].getStaff());
                cell4.SetCellValue(sheetList[i].getAssignment());
                cell5.SetCellValue(sheetList[i].getNotes());
                
                cell.CellStyle = style1;
                cell2.CellStyle = style1;
                spacer.CellStyle = style2;
                cell3.CellStyle = style2;
                cell4.CellStyle = style3;
                cell5.CellStyle = style4;
                noteExtOne.CellStyle = style5;
                noteExtTwo.CellStyle = style5;

                writeRow++;
                
                

            }

            sheet.AutoSizeColumn(5);
            sheet.AutoSizeColumn(6);

            using (FileStream fs = new FileStream(dest, FileMode.Create, FileAccess.Write))
            {
                xssfwb.Write(fs);
            }


        }





        public static bool IsAllUpper(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                if (!Char.IsUpper(input[i]) && input[i] != '-')
                    return false;
            }

            return true;
        }






    }

}
