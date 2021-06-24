using Aspose.Cells;
using Info.Hnbc.InvigilatorExcel.ExcelHandler.ExcelDtos;
using Info.Hnbc.InvigilatorExcel.ExcelHandler.Models;
using Info.Hnbc.Utils.Document.Excel;
using Info.Hnbc.Utils.Document.Excel.Dtos;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Services
{
    public class BianhaoService
    {

        public static KaoshiAnpai Test(string filePath)
        {

            var list = ExcelOperation.ReadExcel<KaoshibianhaoDto>(filePath, null);
            var teachers = ExcelOperation.ReadExcel<TeacherDto>(filePath, null, workSheet: 1);


            var rooms = list.Items.Where(w => !string.IsNullOrWhiteSpace(w.Room)).Select(s => new Room() { Name = s.Room }).ToList();


            var times = list.Items.Where(w => !string.IsNullOrWhiteSpace(w.ExamDate)).Select(s => new KaoshiTime()
            {
                Subject = s.ExamSubject,
                Time = s.ExamDate
            }).ToList();

            var fees = list.Items.Where(w => !string.IsNullOrWhiteSpace(w.TypeForFee)).Select(s => new KaoshiFee()
            {
                Type = s.Type,
                Fee = int.Parse(s.Fee)
            }).ToList();

            var kaoshiinfos = list.Items.Where(w => !string.IsNullOrWhiteSpace(w.ExamNo)).Select(s =>
            {
                var dto = new KaoshiInfo()
                {
                    ExamNo = s.ExamNo,
                    StudentNum = s.StudentNum,
                    Subject = s.Subject,
                    Type = s.Type
                };
                dto.Time = times.FirstOrDefault(f => f.Subject == s.Subject)?.Time;
                dto.Fee = fees.FirstOrDefault(f => f.Type == s.Type)?.Fee ?? 150;

                return dto;
            }
            ).ToList();

            var categories = teachers.Items.Where(w => !string.IsNullOrWhiteSpace(w.Category)).Select(s => new CategoryInfo()
            {
                Category = s.Category,
                Letter = s.Letter
            }).ToList();

            var teacher1 = teachers.Items.Where(w => !string.IsNullOrWhiteSpace(w.Teacher1)).Select(s =>
            {
                var dto = new Teacher()
                {
                    School = s.School1,
                    Name = s.Teacher1,
                    Subject = s.Subject1,
                    Limit = GetCategoryByLetters(s.Limit1, categories),
                    Must = GetCategoryByLetters(s.Must1, categories)
                };
                try
                {
                    dto.LimitNum = int.Parse(s.LimitNum1);
                }
                catch
                {
                    dto.LimitNum = null;
                }
                return dto;
            }).ToList();

            var teacher2 = teachers.Items.Where(w => !string.IsNullOrWhiteSpace(w.Teacher2)).Select(s =>
            {
                var dto = new Teacher()
                {
                    School = s.School2,
                    Name = s.Teacher2,
                    Subject = s.Subject2,
                    Limit = GetCategoryByLetters(s.Limit2, categories),
                    Must = GetCategoryByLetters(s.Must2, categories)
                };
                try
                {
                    dto.LimitNum = int.Parse(s.LimitNum2);
                }
                catch
                {
                    dto.LimitNum = null;
                }
                return dto;
            }).ToList();


            return new KaoshiAnpai
            {
                Teacher1 = teacher1,
                Teacher2 = teacher2,
                Rooms = rooms,
                KaoshiInfos = kaoshiinfos
            };
        }


        public static void Export(KaoshiAnpai ap,string oldFileFullName, string fileFullName)
        { 
            var sheets = new List<ComplexExportDto>();
            var huizong = GetGerenHuizong(ap);
            sheets.Add(huizong);

            var timehuizong = GetTimeTongji(ap);
            sheets.AddRange(timehuizong);

            ExcelOperation.ExportComplexExcelWithOld(sheets.ToArray(), fileFullName, oldFileFullName,2);

        }
        private static List<string> GetCategoryByLetters(string letters, List<CategoryInfo> categories)
        {
            var cletters = letters.ToArray().Select(s => s.ToString()).ToList();
            var cs = categories.Where(w => cletters.Contains(w.Letter))
                .Select(s => s.Category).ToList();

            return cs;
        }

        private static ComplexExportDto GetGerenHuizong(KaoshiAnpai ap)
        {
            var subjects = ap.KaoshiInfos.Select(s => s.Time + "$" + s.Subject).Distinct().OrderBy(o=>o).ToList();


            var cols = new List<ExcelColumnDto>();
            var titleStyle = StyleConst.Header_Right();
            titleStyle.Font.Size = 18;

            cols.Add(new ExcelColumnDto(0, 0, "教师个人监考场次汇总表", 1, 5 + subjects.Count, style: titleStyle));
            cols.Add(new ExcelColumnDto(1, 0, "学校", 2, columnWidth: 8.2, style: StyleConst.Header_Left_Bottom(),isChangeBackground:true));
            cols.Add(new ExcelColumnDto(1, 1, "教师", 2, columnWidth: 8.2, style: StyleConst.Header_Bottom(), isChangeBackground: true));
            cols.Add(new ExcelColumnDto(1, 2, "学科", 2, columnWidth: 8.2, style: StyleConst.Header_Right_Bottom(), isChangeBackground: true));

            for (int i = 0; i < subjects.Count; i++)
            {
                var s = subjects[i];
                var time_sub = s.Split('$');
                var time = time_sub[0];
                var subject = time_sub[1];
                var timeStyle = i == 0 ? StyleConst.Header_Left() : StyleConst.BaseHeaderStyle();
                timeStyle.Font.Name = "华文楷体";
                timeStyle.Font.Size = 14;
                timeStyle.BackgroundColor = Color.FromArgb(255,233,217) ;
                var subStyle = StyleConst.Header_Left_Bottom();
                subStyle.Font.Name = "华文行楷";
                subStyle.Font.Size = 12;
                subStyle.BackgroundColor = Color.FromArgb(255, 233, 217);
                cols.Add(new ExcelColumnDto(1, 3 + i, time, columnWidth: 9, style: timeStyle, isChangeBackground: true));
                cols.Add(new ExcelColumnDto(2, 3 + i, subject, columnWidth: 9, style: subStyle, isChangeBackground: true));
            }

            cols.Add(new ExcelColumnDto(1, 3 + subjects.Count, "监考场数", 2, columnWidth: 11, style: StyleConst.Header_Bottom(), isChangeBackground: true));
            cols.Add(new ExcelColumnDto(1, 4 + subjects.Count, "监考费用", 2, columnWidth: 11, style: StyleConst.Header_Right_Bottom(), isChangeBackground: true));

            List<List<RowData>> data = new List<List<RowData>>();
            int index = 0;
            var teachers = new List<Teacher>();
            teachers.AddRange(ap.Teacher1);
            teachers.AddRange(ap.Teacher2);
            foreach (var l in teachers)
            {
                Style style_1 = StyleConst.Item_Left();
                Style style_2 = StyleConst.BaseItemStyle(); 
                Style style_3 = StyleConst.Item_Right();
                if (index==0)
                {
                    style_1 = StyleConst.Item_Left_Top();
                    style_2 = StyleConst.Item_Top();
                    style_3 = StyleConst.Item_Top_Right();
                }



                var needbackground = index % 2 == 1;
                var dic = new List<RowData>();
                dic.Add(new RowData(0, l.School, style_1, needbackground));
                dic.Add(new RowData(1, l.Name, style_2, needbackground));
                dic.Add(new RowData(2, l.Subject, style_3, needbackground));

                Style cellstyle = StyleConst.BaseItemStyle();

                for(int i=0;i<subjects.Count;i++)
                {
                    var thisstyle = cellstyle;
                    dic.Add(new RowData(3 + i, "", thisstyle, needbackground));
                }


                for (int i =0;i<l.JiankaoSubject.Count;i++)
                {
                    var name = l.JiankaoTime[i] + "$" + l.JiankaoSubject[i] ;
                    var jiankaosub =  l.JiankaoSubject[i]+l.JiankaoType[i];
                    var subindex = subjects.IndexOf(name);

                    var thiscell = dic.Where(w => w.RowIndex == 3 + subindex).FirstOrDefault();
                    thiscell.Content = jiankaosub; 
                }

                dic.Add(new RowData(3 + subjects.Count, l.Jiankao.Count.ToString(), style: StyleConst.BaseItemStyle(),isChangeBackground: needbackground));
                dic.Add(new RowData(4 + subjects.Count, l.JiankaoFee.Sum().ToString(), style: StyleConst.Item_Right(), needbackground));

                data.Add(dic);

                index++;
            }

            
            //添加最后一行封边
            var lastrow = new List<RowData>();
            for(int i=0;i< 5 + subjects.Count;i++)
            {
                lastrow.Add(new RowData(i,"",StyleConst.Item_LastRow()));
            }
            data.Add(lastrow);
            
            return new ComplexExportDto
            {
                Columns = cols,
                Data = data,
                OtherInfo = "个人汇总",
                StartRow = 3
            };
        }



        private static List<ComplexExportDto> GetTimeTongji(KaoshiAnpai ap)
        {
            var result = new List<ComplexExportDto>();
            var timegroups = ap.KaoshiInfos
                .GroupBy(g =>
                g.Time.Replace("上午", "").Replace("下午", "").Replace(".", "月") + "日"
                + g.Type
                )
                .OrderBy(o => o.Key).ToList();

            foreach (var g in timegroups)
            {
                var dto = GetOneDay(g.ToList(), g.Key);
                result.Add(dto);
            }
            return result;
        }


        private static ComplexExportDto GetOneDay(List<KaoshiInfo> ks,string name)
        {
            var subjects = ks.Select(s => s.Subject + "    " + s.Type).Distinct().ToList();


            List<List<RowData>> data = new List<List<RowData>>();
            var cols = new List<ExcelColumnDto>();

            var titleStyle = StyleConst.Header_Right();
            titleStyle.Font.Size = 18;

            cols.Add(new ExcelColumnDto(0, 0, name+"监考安排表", 1, 1+ subjects.Count*6,style: titleStyle));
            cols.Add(new ExcelColumnDto(1, 0, "序号", style: StyleConst.Header_Left_Top())); 

            for (int i = 0; i < subjects.Count; i++)
            {
                var s = subjects[i];
                cols.Add(new ExcelColumnDto(1, 1 + i*6, s, style: StyleConst.Header_Top()));
                cols.Add(new ExcelColumnDto(1, 2 + i*6, "教室", style: StyleConst.Header_Top()));
                cols.Add(new ExcelColumnDto(1, 3 + i*6, "人数", style: StyleConst.Header_Top()));
                cols.Add(new ExcelColumnDto(1, 4 + i*6, "监考甲", style: StyleConst.Header_Top()));
                cols.Add(new ExcelColumnDto(1, 5 + i*6, "监考乙", style: StyleConst.Header_Top()));
                cols.Add(new ExcelColumnDto(1, 6 + i*6, "验收\r\n人员", style: StyleConst.Header_Top_Right()));

                var subjks = ks.Where(w => (w.Subject + "    " + w.Type) == s).ToList();
                int index = 0;
                foreach(var k in subjks)
                {
                    List<RowData> row;
                    if(index>= data.Count)
                    {
                        row = new List<RowData>();
                        row.Add(new RowData(0, index.ToString(), style: StyleConst.BaseItemStyle()));
                        data.Add(row);
                    }
                    else
                    {
                        row = data[index];
                    }

                    var teacher1 = k.Teachers.FirstOrDefault();
                    var teacher2 = k.Teachers.Count > 1 ? k.Teachers.LastOrDefault() : null;
                    row.Add(new RowData(1 + i * 6, k.ExamNo, style: StyleConst.BaseItemStyle()));
                    row.Add(new RowData(2 + i * 6, k.Room?.Name, style: StyleConst.BaseItemStyle()));
                    row.Add(new RowData(3 + i * 6, k.StudentNum, style: StyleConst.BaseItemStyle()));
                    row.Add(new RowData(4 + i * 6, teacher1?.Name, style: StyleConst.BaseItemStyle()));
                    row.Add(new RowData(5 + i * 6, teacher2?.Name, style: StyleConst.BaseItemStyle()));
                    row.Add(new RowData(6 + i * 6, "", style: StyleConst.Item_Right()));
                    index++;
                }

            }

            //添加最后一行封边
            var lastrow = new List<RowData>();
            for (int i = 0; i < 1 + subjects.Count * 6; i++)
            {
                lastrow.Add(new RowData(i, "", StyleConst.Item_LastRow()));
            }
            data.Add(lastrow);

            return new ComplexExportDto
            {
                Columns = cols,
                Data = data,
                OtherInfo = name,
                StartRow = 2
            };
        }
    }
}
