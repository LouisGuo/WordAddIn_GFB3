using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddInPaperCutter.JsonClass
{
    class ExerciseWithRange
    {
        private List<Exercise> _exerciseList;
        private List<Knowlege> _knowlegeList;

        private string _question;
        private string _answer;
        private string _analysis;
        private int _type;
        private int _split;
        private int _score;
        private int _predictDifficult;

        private string _source;
        private int _questionType;
        private int _answerNumber;
        private string _video;
        private string _answerTips;

        private int _rangeStart;
        private int _rangeEnd;




        public ExerciseWithRange()
        {

        }

        public ExerciseWithRange(string question)
        {
            this._question = question;
        }

        public ExerciseWithRange(string question, string answer, string analysis, int type, int split, int score, int predictDifficult)
        {
            this._question = question;
            this._answer = answer;
            this._analysis = analysis;
            this._type = type;
            this._split = split;
            this._score = score;
            this._predictDifficult = predictDifficult;
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

        public List<Knowlege> KnowlegeList
        {
            get
            {
                if (_knowlegeList == null)
                    _knowlegeList = new List<Knowlege>();
                return _knowlegeList;
            }
            set { _knowlegeList = value; }

        }

        public string Question
        {
            get 
            {
                if (_question == null)
                    _question = "";
                return _question; 
            }
            set { _question = value; }
        }

        public string Answer
        {
            get 
            {
                if (_answer == null)
                    _answer = "";
                return _answer; 
            }
            set { _answer = value; }
        }

        public string Analysis
        {
            get 
            {
                if (_analysis == null)
                    _analysis = "";
                return _analysis; 
            }
            set { _analysis = value; }
        }

        public int Type
        {
            get 
            {
                return _type; 
            }
            set { _type = value; }
        }

        public int Split
        {
            get
            {
                return _split;
            }
            set { _split = value; }
        }

        public int Score
        {
            get
            {
                return _score;
            }
            set { _score = value; }
        }

        public int PredictDifficult
        {
            get
            {
                return _predictDifficult;
            }
            set { _predictDifficult = value; }
        }


        public string Source
        {
            get
            {
                if (_source == null)
                    _source = "";
                return _source;
            }
            set
            {
                _source = value;
            }
        }

        public int QuestionType
        {
            get { return _questionType; }
            set { _questionType = value; }
        }

        public int AnswerNumber
        {
            get { return _answerNumber; }
            set { _answerNumber = value; }
        }

        public string Video
        {
            get
            {
                if (_video == null)
                    _video = "";
                return _video;
            }
            set { _video = value; }

        }

        public string AnswerTips
        {
            get
            {
                if (_answerTips == null)
                    _answerTips = "";
                return _answerTips;
            }
            set { _answerTips = value; }
        }

        public int RangeStart
        {
            get { return _rangeStart; }
            set { _rangeStart = value; }
        }

        public int RangeEnd
        {
            get { return _rangeEnd; }
            set { _rangeEnd = value; }
        }
    }
}
