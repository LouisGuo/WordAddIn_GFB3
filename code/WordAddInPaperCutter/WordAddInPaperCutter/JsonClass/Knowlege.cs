using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{
    public class Knowlege
    {
        private int _id;
        private string _name;
        private int _fatherId;


        public int Id
        {
            get
            {
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

        public string FatherId
        {
            get
            {
                return _fatherId.ToString();
            }
            set 
            {
                int output;

                if (int.TryParse(value, out output))
                {
                    _fatherId = int.Parse(value);
                }
            }
        }
    }
}
