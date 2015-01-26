using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordAddInPaperCutter.JsonClass;
using Newtonsoft.Json;

namespace WordAddInPaperCutter.Common
{
    public class JsonFileHelper
    {
        
        public string GetFileString(string filePath)
        {
            string jsonResult = "";
            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close();
            }
            try
            {
                FileStream aFile = new FileStream(filePath, FileMode.OpenOrCreate);
                StreamReader sr = new StreamReader(aFile);
                jsonResult = sr.ReadToEnd();
                sr.Close();
                aFile.Dispose();
            }
            catch
            {

            }
            return jsonResult;
        }
        public ProblemSet GetProblemSetFromFile(string filePath)
        {
            ProblemSet problemSetTemp = new ProblemSet();
            string jsonResult = GetFileString(filePath);
            if ((ProblemSet)JsonConvert.DeserializeObject(jsonResult, typeof(ProblemSet)) !=null)
                problemSetTemp = (ProblemSet)JsonConvert.DeserializeObject(jsonResult, typeof(ProblemSet));
            return problemSetTemp;
        }

        public Paper GetPaperFromFile(string filePath)
        {
            Paper paperTemp = new Paper();
            string jsonResult = GetFileString(filePath);
            if ((Paper)JsonConvert.DeserializeObject(jsonResult, typeof(Paper))!=null)
                paperTemp = (Paper)JsonConvert.DeserializeObject(jsonResult, typeof(Paper));
            return paperTemp;
        }

        public bool WriteFileString(string jsonResult, string filePath)
        {
            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close();
            }
            try
            {
                FileStream aFile = new FileStream(filePath, FileMode.OpenOrCreate);
                aFile.SetLength(0);
                StreamWriter sw = new StreamWriter(aFile);
                sw.WriteLine(jsonResult);
                sw.Close();
                sw.Dispose();
                aFile.Dispose();
                return true;
            }
            catch
            {

            }
            return false;
        }


    }
}
