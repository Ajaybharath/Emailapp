using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Emailapp
{
    public class msg91
    {
        public string country { get; internal set; }
        public string sender { get; internal set; }
        public string route { get; internal set; }
        public List<msg> sms { get; internal set; }
        public string DLT_TE_ID { get; internal set; }
    }
}
