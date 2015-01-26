using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{
    public class PaperNode
    {
        private List<PaperNode> _paperNodeList;
        private List<Exercise> _exerciseList;
        private string _title;

        public PaperNode()
        {

        }

        public PaperNode(string title)
        {
            this._title = title;
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

        public List<Exercise> ExerciseList
        {
            get
            {
                if (_exerciseList == null)
                {
                    _exerciseList = new List<Exercise>();
                }
                return _exerciseList;
            }
            set { _exerciseList = value; }
        }

        public string Title
        {
            get 
            {
                if (_title == null)
                    _title = "";
                return _title; 
            }
            set { _title = value; }
        }
    }
}
