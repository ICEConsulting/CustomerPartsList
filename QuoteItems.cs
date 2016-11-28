using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Tecan_Parts
{
    [Serializable]
    public class QuoteItems
    {
        private string itemSAPID;
        private string itemDescription;
        private int itemQuantity;
        private decimal itemPrice;
        //private decimal itemDiscount;
        //private bool itemNote;
        //private bool itemImage;

        public QuoteItems()
        {
        }

        [XmlAttribute]
        public string SAPID
        {
            get
            {
                return itemSAPID;
            }
            set
            {
                itemSAPID = value;
            }
        }

        [XmlAttribute]
        public string Description
        {
            get
            {
                return itemDescription;
            }
            set
            {
                itemDescription = value;
            }
        }


        [XmlAttribute]
        public int Quantity
        {
            get
            {
                return itemQuantity;
            }
            set
            {
                itemQuantity = value;
            }
        }

        [XmlAttribute]
        public decimal Price
        {
            get
            {
                return itemPrice;
            }
            set
            {
                itemPrice = value;
            }
        }

        //[XmlAttribute]
        //public decimal Discount
        //{
        //    get
        //    {
        //        return itemDiscount;
        //    }
        //    set
        //    {
        //        itemDiscount = value;
        //    }
        //}

        //[XmlAttribute]
        //public bool IncludeNote
        //{
        //    get
        //    {
        //        return itemNote;
        //    }
        //    set
        //    {
        //        itemNote = value;
        //    }
        //}
        //[XmlAttribute]
        //public bool IncludeImage
        //{
        //    get
        //    {
        //        return itemImage;
        //    }
        //    set
        //    {
        //        itemImage = value;
        //    }
        //}
 
    }
}
