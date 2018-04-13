using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;

namespace HubSpotIntegration.Model
{
    public class Company 
    {
        public int company_id { get; set; }
        public string name { get; set; }
        public string website { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string zip { get; set; }
        public string phone { get; set; }
    }
}
