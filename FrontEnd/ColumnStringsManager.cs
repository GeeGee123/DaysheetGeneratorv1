using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrontEnd
{
    class ColumnStringsManager
    {

        private readonly GeneralCSV AdditionalNotes = new GeneralCSV(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 0);
        private readonly GeneralCSV assignToNotes = new GeneralCSV(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 1);
        private readonly GeneralCSV wording = new GeneralCSV(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 2);
        private readonly GeneralCSV notesByName = new GeneralCSV(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 3);
        private readonly GeneralCSV wordsToBold = new GeneralCSV(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 5);

        ResidentLoader resLoader = new ResidentLoader(@"C:\DaySheetGenerator\DaysheetConditions.xlsx", 4);

        List<string> internalAssignmentArray = new List<string>(); // arrAssign list used to contain multiple assignments due to multiple lightning-bolt rows being applied to a single staff member
        List<string> internalNotesArray = new List<string>(); // arrNotes used to contain multiple notes for a single staff member

        public ColumnStringsManager()
        {
            List<string> resList = resLoader.returnList();
        }

        public string PostCall(string postCallInput)
        {
            return postCallInput;
        }

        public string OnCall(string onCallInput, List<string> assignmentArray)
        {
            internalAssignmentArray = assignmentArray;
            if(onCallInput.Trim(' ') == "FC")
                                {
                internalAssignmentArray.Add("PreCall");
            }
            return onCallInput;
        }

        public string StaffName(string staffNameInput, List<string> notesArray)
        {
            internalNotesArray = notesArray;
            string staffNameProcess = staffNameInput;
            staffNameProcess = staffNameProcess.Replace(",", "");

            string[] staffNameOutputmeArray = staffNameProcess.Split(' ');
            string staffNameOutput = "";

            for (int i = 0; i < staffNameOutputmeArray.Length; i++)
            {
                if (staffNameOutputmeArray[i].Length != 0)
                {
                    if (IsAllUpper(staffNameOutputmeArray[i]))
                    {
                        staffNameProcess = staffNameOutputmeArray[i].ToLower();
                        staffNameProcess = (char.ToUpper(staffNameProcess[0])).ToString() + staffNameProcess.Substring(1);
                        staffNameOutput = staffNameOutput + staffNameProcess + " ";

                    }
                }
            }
            staffNameOutput = staffNameOutput.Trim();

            if (notesByName.DictContains(staffNameOutput))
            {
                internalNotesArray.Add(notesByName.getValue(staffNameOutput));
            }

            return staffNameOutput;
        }

        public string Assignment(string assignmentInput)
        {

        }

        public string Notes(string notesInput)
        {

        }

        public List<string> getInternalArray(string arrayName)
        {
            if (arrayName == "assignments")
            {
                return internalAssignmentArray;
            }

            else
            {
                return internalNotesArray;
            }
        }

        private bool IsAllUpper(string input)
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
