using System;
using System.Collections.Generic;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Models
{
    public class KaoshiInfo
    {
        public string ExamNo { get; set; }

        public string Subject { get; set; }

        public string Type { get; set; }

        public string StudentNum { get; set; }

        public string Time { get; set; }

        public int Fee { get; set; }

        public List<Teacher> Teachers { get; set; }
        public Room Room { get; set; }



    }
    public class KaoshiTime
    {
        public string Subject { get; set; }
        public string Time { get; set; }

    }

    public class KaoshiFee
    {
        public string Type { get; set; }
        public int Fee { get; set; }

    }
}
