using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Tecan_Parts
{

    [XmlRootAttribute("Profile", Namespace = "", IsNullable = false)]
    public class Profile
    {
        public Profile()
        { 
        }

        [XmlElementAttribute("Name")]
        public string Name;

        [XmlElementAttribute("Email")]
        public string Email;

        [XmlElementAttribute("Phone")]
        public string Phone;

        [XmlElementAttribute("Company")]
        public string Company;

        [XmlElementAttribute("ShippingAddress")]
        public string ShippingAddress;

        [XmlElementAttribute("City")]
        public string City;

        [XmlElementAttribute("State")]
        public string State;

        [XmlElementAttribute("Zipcode")]
        public string Zipcode;

        [XmlElementAttribute("TecanEmail")]
        public string TecanEmail;

        [XmlElementAttribute("DistributionFolder")]
        public string DistributionFolder;

        [XmlElementAttribute("DatabaseCreationDate")]
        public DateTime DatabaseCreationDate;
    }
}
