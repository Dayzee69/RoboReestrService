using System;

namespace RoboReestrService
{
    class TimerStruct
    {
        public double interval;
        public string format;
        public string condition;

        public TimerStruct(string strinterval, string strformat, string strcondition) 
        {
            
            interval = Double.Parse(strinterval);
            format = strformat;
            condition = strcondition;
        }
    }
}
