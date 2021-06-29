using System;
using System.Collections.Generic;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Models
{
    public class Teacher
    {
        public Teacher()
        {
            Jiankao = new List<string>();
            JiankaoSubject = new List<string>();
            JiankaoSubjectForAnpai = new List<string>();
            JiankaoType = new List<string>();
            JiankaoTime = new List<string>();
            JiankaoFee = new List<int>();
        }

        public string School { get; set; }

        public string Name { get; set; }

        public string Subject { get; set; }

        public int? LimitNum { get; set; }

        public List<SubjectInfo> Limit { get; set; }

        public List<SubjectInfo> Must { get; set; }
        public List<string> Jiankao { get; set; }
        public List<string> JiankaoSubject { get; set; }
        /// <summary>
        /// 仅仅用来安排教师监考使用，不做表格处理
        /// </summary>
        public List<string> JiankaoSubjectForAnpai { get; set; }
        public List<string> JiankaoType { get; set; }
        public List<string> JiankaoTime { get; set; }
        public List<int> JiankaoFee { get; set; }
    }


    public class SubjectInfo
    {
        public string Subject { get; set; }

        public string Letter { get; set; }
    }
}
