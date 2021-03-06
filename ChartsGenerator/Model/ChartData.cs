﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChartsGenerator.Model
{
    public class ChartData
    {
        public string Project { get; set; }
        public string Phase { get; set; }
        public string Task { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Fleet { get; set; }
        public string Color { get; set; }
        public string Vendor { get; set; }
    }

    public class ColorData
    {
        public int No { get; set; }
        public string Task { get; set; }
        public string Color { get; set; }
    }

}