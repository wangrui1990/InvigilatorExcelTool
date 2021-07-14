using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Models
{
    public class KaoshiAnpai
    {
        public List<SubjectInfo> Subjects { get; set; }
        public Dictionary<char,List<string>> LetterSubs { get; set; }


        public List<Teacher> Teacher1 { get; set; }
        public List<Teacher> Teacher2 { get; set; }

        public List<KaoshiInfo> KaoshiInfos { get; set; }

        public List<Room> Rooms { get; set; }

        public List<KaoshiTime> KaoshiTimes { get; set; }


        public void SetJianKao()
        {
            LetterSubs = new Dictionary<char, List<string>>();
            foreach(var s in Subjects)
            {
                if(s.Letter.Length<=1)
                {
                    continue;
                }
                var letter = s.Letter[0];
                var list = new List<string>();
                if(LetterSubs.ContainsKey(letter))
                {
                    list = LetterSubs[letter];
                    list.Add(s.Subject);
                }
                else
                {
                    list.Add(s.Subject);
                    LetterSubs.Add(letter,list);
                }
            }

            //统计场数，用于教师合理安排监考，有场数限制的教师优先安排到场数多的考试，
            foreach (var k in KaoshiInfos)
            {
                k.SubjectCount = KaoshiInfos.Count(c => c.Subject == k.Subject);
                k.SubjectLetter = Subjects.FirstOrDefault(f=>f.Subject == k.Subject)?.Letter;
            }

            //场数多的考试优先安排，以防止最后因部分老师有监考场数限制而导致监考老师不足
            var kslist = KaoshiInfos.OrderByDescending(o => o.SubjectCount).ToList();
            var ddd = kslist.Where(w=>w.Subject=="地理").ToList();
            foreach (var k in kslist)
            {
                if (k.Subject == "技术")
                {

                }
                k.Teachers = new List<Teacher>();

                Teacher1.ForEach(f => {
                    f.Order = "";
                    f.Order += (f.Limit.Count(c => c.Subject == k.Subject)>0) ? "01" : "00";
                    f.Order += (f.Must.Count(c => c.Subject == k.Subject)>0) ? "00" : "01";
                    f.Order += (f.Must.Count(c => c.Subject != k.Subject) > 0)?"01":"00";
                    f.Order += (f.Subject == k.Subject) ?"01":"00";
                    f.Order += (f.JiankaoTime.Contains(k.Time)) ?"00":"01";
                    f.Order += f.Jiankao.Count.ToString().PadLeft(3,'0');
                });

                var t1 = Teacher1
                     .Where(w => !w.JiankaoSubjectForAnpai.Contains(k.Subject))
                     .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value)) //排除掉超过监考限制的教师
                    .OrderBy(o=>o.Order)
                    .FirstOrDefault();

                //var t1 = Teacher1
                //    .Where(w => !w.JiankaoSubjectForAnpai.Contains(k.Subject)) //已经安排了同一科目考试监考的教师排除
                //    .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value)) //排除掉超过监考限制的教师
                //    .OrderByDescending(o => o.Must.Count(c=>c.Subject == k.Subject)) //优先选择设置了必须监考的教师
                //    .ThenBy(o => o.Must.Count(c => c.Subject != k.Subject) > 0)//尽量避免已经设置了必须监考的教师，以保证该教师能监考到必须监考的科目
                //    .ThenBy(w => w.Subject == k.Subject) //教师自己教学的科目后安排
                //    .ThenByDescending(o => o.JiankaoTime.Contains(k.Time)) //监考老师尽量的安排同一上午或同一下午
                //    .ThenBy(o => o.Jiankao.Count) //优先选择没有安排过监考的或者监考数量最少的
                //    .ThenBy(o => o.Limit.Count(c => c.Subject == k.Subject)) //设置了避免监考的最后考虑
                //    .FirstOrDefault();


                if (t1 != null)
                {
                    k.Teachers.Add(t1);
                    t1.Jiankao.Add(k.ExamNo);

                    t1.JiankaoSubject.Add(k.Subject);
                    if (k.SubjectLetter.Length > 1)
                    {
                        t1.JiankaoSubjectForAnpai.AddRange(LetterSubs[k.SubjectLetter[0]]);
                    }
                    else
                    {
                        t1.JiankaoSubjectForAnpai.Add(k.Subject);
                    }
                    t1.JiankaoType.Add(k.Type);
                    t1.JiankaoTime.Add(k.Time);
                    t1.JiankaoFee.Add(k.Fee);
                }



                Teacher2.ForEach(f => {
                    f.Order = "";
                    f.Order += (f.Limit.Count(c => c.Subject == k.Subject) > 0) ? "01" : "00";
                    f.Order += (f.Must.Count(c => c.Subject == k.Subject) > 0) ? "00" : "01";
                    f.Order += (f.Must.Count(c => c.Subject != k.Subject) > 0) ? "01" : "00";
                    f.Order += (f.Subject == k.Subject) ? "01" : "00";
                    f.Order += (f.JiankaoTime.Contains(k.Time)) ? "00" : "01";
                    f.Order += f.Jiankao.Count.ToString().PadLeft(3, '0');
                });

                var tt2 = Teacher2
                     .Where(w => !w.JiankaoSubjectForAnpai.Contains(k.Subject))
                     .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value)) //排除掉超过监考限制的教师
                    .OrderBy(o => o.Order)
                    .ToList();
                var t2 = Teacher2
                     .Where(w => !w.JiankaoSubjectForAnpai.Contains(k.Subject))
                     .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value)) //排除掉超过监考限制的教师
                    .OrderBy(o => o.Order)
                    .FirstOrDefault();

                //var t2 = Teacher2
                //    .Where(w => !w.JiankaoSubjectForAnpai.Contains(k.Subject))
                //    .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value))
                //    .OrderByDescending(o => o.Must.Count(c => c.Subject == k.Subject))
                //    .ThenBy(o => o.Must.Count(c => c.Subject != k.Subject) > 0)
                //    .ThenBy(w => w.Subject == k.Subject)  
                //    .ThenByDescending(o => o.JiankaoTime.Contains(k.Time))
                //    .ThenBy(o => o.Jiankao.Count)
                //    .ThenBy(o => o.Limit.Count(c => c.Subject == k.Subject))
                //    .FirstOrDefault();
                if (t2 != null)
                {
                    k.Teachers.Add(t2);
                    t2.Jiankao.Add(k.ExamNo);
                    t2.JiankaoSubject.Add(k.Subject);
                    if (k.SubjectLetter.Length > 1)
                    {
                        //如果是英语或日语等同时考试的科目，需要将其同
                        t2.JiankaoSubjectForAnpai.AddRange(LetterSubs[k.SubjectLetter[0]]);
                    }
                    else
                    {
                        t2.JiankaoSubjectForAnpai.Add(k.Subject);
                    }
                    t2.JiankaoType.Add(k.Type);
                    t2.JiankaoTime.Add(k.Time);
                    t2.JiankaoFee.Add(k.Fee);
                }


                var r = Rooms.Where(o => !o.Zhanyong.Contains(k.Subject))
                    //.OrderBy(o => o.Zhanyong.Count) //设置使用率低的教室优先
                    .FirstOrDefault();

                k.Room = r;
                if (r != null)
                {
                    r.Zhanyong.Add(k.Subject);
                }
            }
        }
    }
}
