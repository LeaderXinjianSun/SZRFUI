﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZRFUI.Models
{
    public class AlarmData
    {
        public string Code { set; get; }
        public string Content { set; get; }
        public string Type { set; get; }
        public DateTime Start { set; get; }
        public DateTime End { set; get; }
        public bool State { set; get; }
    }
}
