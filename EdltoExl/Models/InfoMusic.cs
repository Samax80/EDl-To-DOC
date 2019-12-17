using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EdltoExl.Models
{
    public class InfoMusic
    {
        public string Music_Name { get; set; }
        public string Music_FullName { get; set; }
        public string Music_Cd { get; set; }
        public string Music_Cue { get; set; }
        public TimeSpan Music_Tc_InTime { get; set; }
        public TimeSpan Music_Tc_OutTime { get; set; }
        public TimeSpan Music_Tc_Duration { get; set; }
        public string Music_Publisher { get; set; }

    }
}