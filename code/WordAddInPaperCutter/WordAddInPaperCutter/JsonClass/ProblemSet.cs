using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{
    public class ProblemSet
    {
        private List<Exercise> _exerciseList;

        private string _name;
        private int _uploadUser;
        private int _id;
        private int _problemSetTypeId;

        public ProblemSet()
        {

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

        public int ProblemSetTypeId
        {
            get { return _problemSetTypeId; }
            set { _problemSetTypeId = value; }
        }
    }
}
