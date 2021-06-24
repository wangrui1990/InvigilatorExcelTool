using Info.Hnbc.Utils.Document.Excel.Attributes;
using Info.Hnbc.Utils.Document.Excel.Dtos;
using System;
using System.Collections.Generic;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.ExcelDtos
{
    public class KaoshibianhaoDto : ImportBaseDto
    {
        public override void Check(object obj = null)
        {
            //throw new NotImplementedException();
        }

        [ExcelColumnName("考场号")]
        public string ExamNo { get; set; }

        [ExcelColumnName("科目")]
        public string Subject { get; set; }

        [ExcelColumnName("考试种类")]
        public string Type { get; set; }

        [ExcelColumnName("人数30")]
        public string StudentNum { get; set; }

        [ExcelColumnName("教室")]
        public string Room { get; set; }


        [ExcelColumnName("考试日期")]
        public string ExamDate { get; set; }

        [ExcelColumnName("考试科目")]
        public string ExamSubject { get; set; }


        [ExcelColumnName("预设考试种类")]
        public string TypeForFee { get; set; }

        [ExcelColumnName("监考费/场")]
        public string Fee { get; set; }


    }
}
