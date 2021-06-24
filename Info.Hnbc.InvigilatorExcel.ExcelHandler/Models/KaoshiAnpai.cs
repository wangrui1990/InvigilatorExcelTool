using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Models
{
    public class KaoshiAnpai
    {

        public List<Teacher> Teacher1 { get; set; }
        public List<Teacher> Teacher2 { get; set; }

        public List<KaoshiInfo> KaoshiInfos { get; set; }

        public List<Room> Rooms { get; set; }


        public void SetJianKao()
        {
            foreach(var k in KaoshiInfos)
            {
                k.Teachers = new List<Teacher>();
                var t1 = Teacher1
                    .Where(w=>!w.JiankaoSubject.Contains(k.Subject)) //已经安排了同一科目考试监考的教师排除
                    .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count>=w.LimitNum.Value)) //排除掉超过监考限制的教师
                    .OrderByDescending(o => o.Must.Contains(k.Subject)) //优先选择设置了必须监考的教师
                    .ThenBy(o => o.Limit.Contains(k.Subject)) //设置了避免监考的最后考虑
                    .ThenBy(o => o.Jiankao.Count) //优先选择没有安排过监考的或者监考数量最少的
                    .ThenByDescending(o=>o.JiankaoTime.Contains(k.Time)) //监考老师尽量的安排同一上午或同一下午
                    .FirstOrDefault();
                if (t1 != null)
                {
                    k.Teachers.Add(t1);
                    t1.Jiankao.Add(k.ExamNo);
                    t1.JiankaoSubject.Add(k.Subject);
                    t1.JiankaoType.Add(k.Type);
                    t1.JiankaoTime.Add(k.Time);
                    t1.JiankaoFee.Add(k.Fee);
                }

                var t2 = Teacher2
                    .Where(w => !w.JiankaoSubject.Contains(k.Subject))
                    .Where(w => !(w.LimitNum.HasValue && w.Jiankao.Count >= w.LimitNum.Value))
                    .OrderByDescending(o => o.Must.Contains(k.Subject))
                    .ThenBy(o => o.Limit.Contains(k.Subject))
                    .ThenBy(o => o.Jiankao.Count)
                    .ThenByDescending(o => o.JiankaoTime.Contains(k.Time))
                    .FirstOrDefault();
                if (t2 != null)
                {
                    k.Teachers.Add(t2);
                    t2.Jiankao.Add(k.ExamNo);
                    t2.JiankaoSubject.Add(k.Subject);
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
