using System;
using System.Collections.Generic;
using System.Text;

namespace FrontEnd
{
    class SheetLine
    {

        private string post;
        private string onCall;
        private string staff;
        private string assignment;
        private string notes;
        private string pType;

        public SheetLine()
        {
            this.post = "";
            this.onCall = "";
            this.staff = "";
            this.assignment = "";
            this.notes = "";
            this.pType = "";

        }

        public void setPost(string post)
        {
            this.post = post;
        }

        public void setOnCall(string onCall)
        {
            this.onCall = onCall;
        }

        public void setStaff(string staff)
        {
            this.staff = staff;
        }

        public void setAssignment(string assignment)
        {
            this.assignment = assignment;
        }

        public void setNotes(string notes)
        {
            this.notes = notes;
        }

        public void setPType(string pType)
        {
            this.pType = pType;
        }

        public string getPost()
        {
            return post;
        }

        public string getOnCall()
        {
            return onCall;
        }

        public string getStaff()
        {
            return staff;
        }

        public string getAssignment()
        {
            return assignment;
        }

        public string getNotes()
        {
            return notes;
        }

        public string getPType()
        {
            return pType;
        }

    }
}