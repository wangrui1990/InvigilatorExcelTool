using Info.Hnbc.Utils.Document.Excel.Attributes;
using Info.Hnbc.Utils.Document.Excel.Dtos;
using System;
using System.Collections.Generic;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.ExcelDtos
{
    public class TeacherDto : ImportBaseDto
    {
        public override void Check(object obj = null)
        {
            //throw new NotImplementedException();
        }

        [ExcelColumnName("学校甲")]
        public string School1 { get; set; }

        [ExcelColumnName("监考甲")]
        public string Teacher1 { get; set; }

        [ExcelColumnName("学科")]
        public string Subject1 { get; set; }

        [ExcelColumnName("场数限制甲")]
        public string LimitNum1 { get; set; }

        [ExcelColumnName("限制监考甲")]
        public string Limit1 { get; set; }

        [ExcelColumnName("必须监考甲")]
        public string Must1 { get; set; }


        [ExcelColumnName("学校乙")]
        public string School2 { get; set; }


        [ExcelColumnName("监考乙")]
        public string Teacher2 { get; set; }


        [ExcelColumnName("学科乙")]
        public string Subject2 { get; set; }



        [ExcelColumnName("场数限制乙")]
        public string LimitNum2 { get; set; }



        [ExcelColumnName("限制监考乙")]
        public string Limit2 { get; set; }



        [ExcelColumnName("必须监考乙")]
        public string Must2 { get; set; }



        [ExcelColumnName("考试科目")]
        public string Category { get; set; }

        [ExcelColumnName("字母化")]
        public string Letter { get; set; }
    }

}
