using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Tecan_Parts
{

    [XmlRootAttribute("Quote", Namespace = "", IsNullable = false)]
    public class Quote
    {
        public Quote()
        {
        }

        [XmlElementAttribute("Title")]
        public string QuoteTitle;
        public DateTime QuoteDate;
        public string QuoteEmailTo;
        public decimal QuoteTotal;

        // Serializes an ArrayList as a "Items" array of XML elements of custom type QuoteItems named "Items".
        [XmlArray("Items"), XmlArrayItem("Item", typeof(QuoteItems))]
        public System.Collections.ArrayList Items = new System.Collections.ArrayList();

        //// Serializes an ArrayList as a "Options" array of XML elements of custom type QuoteItems named "Options".
        //[XmlArray("Options"), XmlArrayItem("Item", typeof(QuoteItems))]
        //public System.Collections.ArrayList Options = new System.Collections.ArrayList();

        //// Serializes an ArrayList as a "ThirdParty" array of XML elements of custom type QuoteItems named "ThirdParty".
        //[XmlArray("ThirdParty"), XmlArrayItem("Item", typeof(QuoteItems))]
        //public System.Collections.ArrayList ThirdParty = new System.Collections.ArrayList();

        //// Serializes an ArrayList as a "SmartStart" array of XML elements of custom type QuoteItems named "SmartStart".
        //[XmlArray("SmartStart"), XmlArrayItem("Item", typeof(QuoteItems))]
        //public System.Collections.ArrayList SmartStart = new System.Collections.ArrayList();

    }
}
