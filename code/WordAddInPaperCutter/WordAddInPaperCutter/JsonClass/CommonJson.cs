using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{

    public class CommonJson
    {
        private int _id;
        private string _name;

        List<CommonJson> _commonJsonList=new List<CommonJson>();


        public List<CommonJson> CommonJsonList
        {
            get
            {
                if (_commonJsonList == null)
                {
                    _commonJsonList = new List<CommonJson>();
                }
                return _commonJsonList;
            }
            set { _commonJsonList = value; }
        }

        public int Id
        {
            get
            {
                if (_id == null)
                    _id = new int();
                return _id;
            }
            set { _id = value; }
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
    }
}
