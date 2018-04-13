using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HubSpotIntegration.Model
{
    public class Contact
    {
        public int vid;
        public string firstname;
        public string lastname;
        public string lifecyclestage;
        public Company associated_company;
    }
}
