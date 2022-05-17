using System;
using System.IO;
using System.Collections.Generic;
using System.Data.Odbc;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;
using System.Text.RegularExpressions;
using MessageBox = System.Windows.MessageBox;

namespace FrontEnd
{
    class Program
    {


        public Program(List<string> filePaths, string destination)
        {
           
            string layout = @"C:\DaySheetGenerator\Layout_Template.xlsx"; // Defining the path to the template the daysheet will be written into
            DateTime dateFile; // DateTime variable that will be used to assign date to file name
            ColumnStringsManager strManager = new ColumnStringsManager();

            //Opens Lighting-Bolt spreadsheet to extract sheet date for use in final file name
            /*****************************************************************************************************************************************/
            using (FileStream file = new FileStream(filePaths[0], FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfwb = new HSSFWorkbook(file);
                
                ISheet sheet = hssfwb.GetSheetAt(0);
                string destDate = sheet.GetRow(2).GetCell(6).ToString();
                dateFile = DateTime.Parse(destDate);
                file.Close();
                file.Dispose();
            }
           /*****************************************************************************************************************************************/
            

            //Creates string for path to output file and output file name
            /****************************************************************************************************************************************/
            string dest = destination + "\\DaySheet_" + dateFile.ToString("MMM") + "_" + dateFile.ToString("dd") + ".xlsx"; // File destination string

            if (File.Exists(dest)) // Checks if file already exists, deletes file if exists
            {
                File.Delete(dest);
            }

            File.Copy(layout, dest); // Copies the layout .xlsx to the destination path
            /****************************************************************************************************************************************/


            

            for (int reportNum = 0; reportNum < filePaths.Count; reportNum++)
            {
                
                List<string> arrAssign = new List<string>(); // arrAssign list used to contain multiple assignments due to multiple lightning-bolt rows being applied to a single staff member
                List<string> arrNotes = new List<string>(); // arrNotes used to contain multiple notes for a single staff member
                List<SheetLine> sheetList = new List<SheetLine>(); // stores object representaion of row data from lightning-bolt XLS

                HSSFWorkbook hssfwb; // initialize excel XLS reader
                XSSFWorkbook xssfwb; // initialize excel XLSX reader


                // Opens XLS file to be read from lightning-bolt spreadsheet input 
                /**************************************************************************************************************************************/
                using (FileStream file = new FileStream(filePaths[reportNum], FileMode.Open, FileAccess.Read)) 
                {
                    hssfwb = new HSSFWorkbook(file);
                    file.Close();
                    file.Dispose();

                }
                /**************************************************************************************************************************************/


                // Management of date representations for the output sheet
                /**************************************************************************************************************************************/
                ISheet sheet = hssfwb.GetSheetAt(0); // Get Report Sheet from xlsx
                SheetLine sheetLine;
                string destDate = sheet.GetRow(2).GetCell(6).ToString();
                DateTime date = DateTime.Parse(destDate);
                /**************************************************************************************************************************************/

                for (int row = 1; row < sheet.PhysicalNumberOfRows - 1; row++)
                {

                    sheetLine = new SheetLine(); // Creates new sheetLine object for eacg row in the spreadsheet


                    /**************************************************************************************************************************************/
                    if ((row < sheet.LastRowNum) && sheet.GetRow(row).GetCell(2).ToString() == sheet.GetRow(row + 1).GetCell(2).ToString()) // Checks if the name of the staff member of the current row is also in the next row
                    {                                                                                                                       // If true, condition adds the assignment and notes of the current row to the arrAssign
                        arrAssign.Add(sheet.GetRow(row).GetCell(3).ToString().Trim());                                                             // and arrNotes Lists respectively, and does not add the line to the pdf
                        arrNotes.Add(sheet.GetRow(row).GetCell(4).ToString().Trim());
                    }
                    /**************************************************************************************************************************************/


                    else
                    {

                        for (int column = 0; column < 6; column++) // Loops the the 5 columns of relevent data in the spreadsheet
                        {

                            // Manages post data from lightning-bolt XLS
                            if (column == 0)
                            {
                                string postCallInput = sheet.GetRow(row).GetCell(column).ToString();
                                string postCallOutput = strManager.PostCall(postCallInput);
                                sheetLine.setPost(postCallOutput); 
                            }

                            // Manages call data from lightning-bolt XLS and applies changes to call data strings
                            if (column == 1)
                            {
                                string onCallInput = sheet.GetRow(row).GetCell(column).ToString();
                                string onCallOutput = strManager.OnCall(onCallInput, arrAssign);
                                arrAssign = strManager.getInternalArray("assignments");
                                sheetLine.setOnCall(onCallOutput); // assigns the 1st column data to sheetline->OnCall
                            }
                             
                            // Manages staff name data from lightning-bolt XLS and applies changes to staff data strings
                            if (column == 2)
                            {
                                string staffNameInput = sheet.GetRow(row).GetCell(column).ToString();
                                string staffNameOutput = strManager.StaffName(staffNameInput, arrNotes);
                                arrNotes = strManager.getInternalArray("notes");
                                sheetLine.setStaff(staffNameOutput);
                            }
                            
                            // Manages staff assignment data from lightning-bolt XLS and applies changes to assignment data strings
                            if (column == 3)
                            {
                                string assignmentInput = sheet.GetRow(row).GetCell(column).ToString();
                                string assignmentOutput = strManager.Assignment(assignmentInput, arrAssign);
                                arrNotes = strManager.getInternalArray("notes");
                                sheetLine.setAssignment(assignmentOutput);
                            }

                            // Manages staff notes data from lightning-bolt XLS and applies changes to notes data strings
                            if (column == 4)
                            {
                                string notesInput = sheet.GetRow(row).GetCell(column).ToString();
                                string notesOutput = strManager.Notes(notesInput);
                                sheetLine.setNotes(notesOutput);
                            }

                            // Stores staff type (Staff, Fellow, AA, Resident etc...) for later use in formatting logic
                            if (column == 5)
                            {
                                sheetLine.setPType(sheet.GetRow(row).GetCell(column).ToString());
                            }

                        }

                        sheetList.Add(sheetLine); // Adds Row representaion sheetLine to sheetList
                        arrAssign.Clear(); // Clears assignment array
                        arrNotes.Clear(); // Clears notes array
                    }


                }

                hssfwb.Close(); // Closes lightning-bolt XLS file 

               
                using (FileStream file = new FileStream(dest, FileMode.Open, FileAccess.ReadWrite)) // Opening spreadsheet with spreadsheet reader 
                {
                    xssfwb = new XSSFWorkbook(file);

                }


                xssfwb.SetSheetName(reportNum, date.ToString("dddd") + " " + date.ToString("MMM") + " " + date.ToString("dd"));
                sheet = xssfwb.GetSheetAt(reportNum);

                //CELL FONT CONTROLS
                /********************************************************************************************************************************************************************/
                var boldFont = xssfwb.CreateFont();
                boldFont.IsBold = true;
                boldFont.FontHeightInPoints = 11;

                var smallBoldFont = xssfwb.CreateFont();
                smallBoldFont.IsBold = true;
                smallBoldFont.FontHeightInPoints = 8;

                var dateFont = xssfwb.CreateFont();
                dateFont.IsBold = true;
                dateFont.FontHeightInPoints = 18;

                var generalFont = xssfwb.CreateFont();
                generalFont.FontHeightInPoints = 11;

                var smallFont = xssfwb.CreateFont();
                smallFont.FontHeightInPoints = 8;
                /********************************************************************************************************************************************************************/

                //CELL STYLE CONTROLS
                /********************************************************************************************************************************************************************/
                var postCellStyle = xssfwb.CreateCellStyle();
                postCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                postCellStyle.BorderBottom = BorderStyle.Thin;
                postCellStyle.SetFont(generalFont);

                var callCellStyle = xssfwb.CreateCellStyle();
                callCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                callCellStyle.BorderBottom = BorderStyle.Thin;
                callCellStyle.SetFont(boldFont);

                var callCellStyle2 = xssfwb.CreateCellStyle();
                callCellStyle2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                callCellStyle2.BorderBottom = BorderStyle.Thin;
                callCellStyle2.SetFont(smallBoldFont);

                var ECellStyle = xssfwb.CreateCellStyle();
                ECellStyle.BorderLeft = BorderStyle.Thin;

                var nameCellStyle = xssfwb.CreateCellStyle();
                nameCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                nameCellStyle.BorderLeft = BorderStyle.Thin;
                nameCellStyle.BorderRight = BorderStyle.Thin;
                nameCellStyle.SetFont(generalFont);

                var assignCellStyle = xssfwb.CreateCellStyle();
                assignCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                assignCellStyle.BorderLeft = BorderStyle.Thin;
                assignCellStyle.BorderRight = BorderStyle.Thin;
                assignCellStyle.BorderBottom = BorderStyle.Thin;
                assignCellStyle.SetFont(generalFont);

                var assignCellStyle2 = xssfwb.CreateCellStyle();
                assignCellStyle2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                assignCellStyle2.BorderLeft = BorderStyle.Thin;
                assignCellStyle2.BorderRight = BorderStyle.Thin;
                assignCellStyle2.BorderBottom = BorderStyle.Thin;
                assignCellStyle2.SetFont(smallFont);

                var notesCellStyle = xssfwb.CreateCellStyle();
                notesCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                notesCellStyle.BorderLeft = BorderStyle.Thin;
                notesCellStyle.BorderBottom = BorderStyle.Dotted;
                notesCellStyle.SetFont(generalFont);

                var notesBoldStyle = xssfwb.CreateCellStyle();
                notesBoldStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                notesBoldStyle.BorderLeft = BorderStyle.Thin;
                notesBoldStyle.BorderBottom = BorderStyle.Dotted;
                notesBoldStyle.SetFont(boldFont);

                var IJCellStyle = xssfwb.CreateCellStyle();
                IJCellStyle.BorderBottom = BorderStyle.Dotted;

                var dateStyle = xssfwb.CreateCellStyle();
                dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                dateStyle.SetFont(dateFont);
                /*******************************************************************************************************************************************************************/

                bool assocFlag = true;
                bool assistFlag = true;
                bool fellowFlag = true;

                List<SheetLine> unorderedRes = new List<SheetLine>();
               
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

                    if (!sheetList[i].getPType().Equals("Resident"))
                    {
                        IRow dataRow = sheet.CreateRow(writeRow);
                        dataRow.HeightInPoints = 14.25F;


                        ICell postCell = dataRow.CreateCell(0);
                        ICell onCallCell = dataRow.CreateCell(2);
                        ICell spacer = dataRow.CreateCell(4);
                        ICell staffCell = dataRow.CreateCell(5);
                        ICell assignCell = dataRow.CreateCell(6);
                        ICell notesCell = dataRow.CreateCell(7);
                        ICell noteExtOne = dataRow.CreateCell(8);
                        ICell noteExtTwo = dataRow.CreateCell(9);



                        string assignTrimmer = sheetList[i].getAssignment().Trim(' ', '\n');
                        string noteTrimmer = sheetList[i].getNotes().Trim(' ', '\n');

                        string[] stringParse = sheetList[i].getStaff().Split(' ');
                        string finalName = stringParse[0].Trim(',') + ", " + stringParse[1].Substring(0, 1);

                        postCell.SetCellValue(sheetList[i].getPost());
                        onCallCell.SetCellValue(sheetList[i].getOnCall());
                        staffCell.SetCellValue(finalName);
                        assignCell.SetCellValue(assignTrimmer);
                        notesCell.SetCellValue(noteTrimmer);

                        spacer.CellStyle = ECellStyle;
                        staffCell.CellStyle = nameCellStyle;
                        postCell.CellStyle = postCellStyle;

                        if (sheetList[i].getOnCall().Length < 10)
                        {
                            onCallCell.CellStyle = callCellStyle;
                        }
                        else
                        {
                            onCallCell.CellStyle = callCellStyle2;
                        }
                        
                        


                        if (assignTrimmer.Length < 20)
                        {
                            assignCell.CellStyle = assignCellStyle;
                        }
                        else
                        {
                            assignCell.CellStyle = assignCellStyle2;
                        }
                        

                        if (wordsToBold.DictContains(noteTrimmer))
                        {
                            notesCell.CellStyle = notesBoldStyle;
                        }
                        else
                        {
                            notesCell.CellStyle = notesCellStyle;
                        }

                        noteExtOne.CellStyle = IJCellStyle;
                        noteExtTwo.CellStyle = IJCellStyle;

                        writeRow++;
                    }
                    else
                    {
                        unorderedRes.Add(sheetList[i]);
                        writeRow++;
                    }

                }

                int resRow = 105;

                for (int i = 0; i < resList.Count(); i++)
                {
                    for (int j = 0; j < unorderedRes.Count(); j++)
                    {

                        if (resList[i] == unorderedRes[j].getStaff())
                        {
                            IRow dataRow = sheet.CreateRow(resRow);
                            dataRow.HeightInPoints = 14.25F;


                            ICell postCell = dataRow.CreateCell(0);
                            ICell onCallCell = dataRow.CreateCell(2);
                            ICell spacer = dataRow.CreateCell(4);
                            ICell staffCell = dataRow.CreateCell(5);
                            ICell assignCell = dataRow.CreateCell(6);
                            ICell notesCell = dataRow.CreateCell(7);
                            ICell noteExtOne = dataRow.CreateCell(8);
                            ICell noteExtTwo = dataRow.CreateCell(9);



                            string assignTrimmer = unorderedRes[j].getAssignment().Trim(' ', '\n');
                            string noteTrimmer = unorderedRes[j].getNotes().Trim(' ', '\n');

                            
                     

                            string[] stringParse = unorderedRes[j].getStaff().Split(' ');
                            string finalName = stringParse[0].Trim(',') + ", " + stringParse[1].Substring(0, 1);



                            postCell.SetCellValue(unorderedRes[j].getPost());
                            onCallCell.SetCellValue(unorderedRes[j].getOnCall());
                            staffCell.SetCellValue(finalName);
                            assignCell.SetCellValue(assignTrimmer);
                            notesCell.SetCellValue(noteTrimmer);

                            postCell.CellStyle = postCellStyle;
                            onCallCell.CellStyle = callCellStyle;
                            spacer.CellStyle = ECellStyle;
                            staffCell.CellStyle = nameCellStyle;
                            if (assignTrimmer.Length < 20)
                            {
                                assignCell.CellStyle = assignCellStyle;
                            }
                            else
                            {
                                assignCell.CellStyle = assignCellStyle2;
                            }

                            if (wordsToBold.DictContains(noteTrimmer))
                            {
                                notesCell.CellStyle = notesBoldStyle;
                            }
                            else
                            {
                                notesCell.CellStyle = notesCellStyle;
                            }

                            noteExtOne.CellStyle = IJCellStyle;
                            noteExtTwo.CellStyle = IJCellStyle;

                            resRow++;
                        }
                    }
                }

                
                sheet.SetColumnWidth(5, 5120);
                sheet.SetColumnWidth(6, 5376);
                using (FileStream fs = new FileStream(dest, FileMode.Create, FileAccess.Write))
                {
                    xssfwb.Write(fs);
                }
            }
        }
    }
}
