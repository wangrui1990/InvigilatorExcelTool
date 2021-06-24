using System;
using System.Collections.Generic;
using System.Text;

namespace Info.Hnbc.InvigilatorExcel.ExcelHandler.Models
{
    public class Room
    {
        public Room()
        {
            Zhanyong = new List<string>();
        }
        public string Name { get; set; }

        public List<string> Zhanyong { get; set; }

    }
}
