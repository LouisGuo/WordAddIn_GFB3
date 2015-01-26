using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{
    public class Paper
    {
        private List<PaperNode> _paperNodeList;


        private string _name;
        private int _uploadUser;
        private int _id;
        private int _paperTypeId;

        public Paper()
        {

        }

        public Paper(string name, int uploadUser)
        {
            this._name = name;
            this._uploadUser = uploadUser;
        }

        public List<PaperNode> PaperNodeList
        {
            get
            {
                if (_paperNodeList == null)
                {
                    _paperNodeList = new List<PaperNode>();
                }
                return _paperNodeList;
            }
            set { _paperNodeList = value; }
        }

        public string Name
        {
            get 
            {
                if (_name == null)
                    _name = "";
                return _name; 
            }
            set { _name = value; }
        }

        public int UploadUser
        {
            get 
            {
                if (_uploadUser == null)
                    _uploadUser = new int();
                return _uploadUser;
            }
            set { _uploadUser = value; }
        }

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int PaperTypeId
        {
            get { return _paperTypeId; }
            set { _paperTypeId = value; }
        }
    }
}
