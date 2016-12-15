using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlServerCe;
using System.Diagnostics;

namespace Tecan_Parts
{
    public partial class PartsListDetailDisplay : Form
    {

        MainQuoteForm mainForm;        
        SqlCeConnection TecanDatabase = null;

        public void SetForm1Instance(MainQuoteForm inst)
        {
            mainForm = inst;
        }

        public PartsListDetailDisplay()
        {
            InitializeComponent();
        }

        internal void LoadParts(String SAPID)
        {
            this.partsListTableAdapter.FillBySAPID(this.tecanCustomerPartsListDataSet.PartsList, SAPID);
            this.suppumentalDocsTableAdapter.FillBySAPID(this.tecanCustomerPartsListDataSet.SuppumentalDocs, SAPID);
            //// this.requiredPartsTableAdapter.FillBySAPID(this.TecanCustomerPartsListDataSet.RequiredParts, SAPID);
            // this.sSPCategoryTableAdapter.Fill(this.tecanQuoteGeneratorPartsListDataSet.SSPCategory);
            // this.salesTypeTableAdapter.Fill(this.tecanQuoteGeneratorPartsListDataSet.SalesType);
            this.subCategoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.SubCategory);
            this.categoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Category);
            this.instrumentTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Instrument);

            setLookupItemsText(SAPID);
            loadImage(SAPID);
            loadRequiredParts(SAPID);
            loadCompatibility(SAPID);
            loadPartsListLookupComboBox();
            sAPIdTextBox.Focus();
        }

        private void setLookupItemsText(String SAPID)
        {
            openDB();
            SqlCeCommand cmd = TecanDatabase.CreateCommand();
            SqlCeDataReader reader;

            // Instrument
            cmd.CommandText = "SELECT P.Instrument, C.InstrumentName FROM PartsList P" +
            " INNER JOIN Instrument C" +
            " ON P.Instrument = C.InstrumentID" +
            " WHERE SAPId = '" + SAPID + "'";

            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                InstrumentTextBox.Text = reader[1].ToString();
            }

            // Category
            cmd.CommandText = "SELECT P.Category, C.CategoryName FROM PartsList P" +
            " INNER JOIN Category C" +
            " ON P.Category = C.CategoryID" +
            " WHERE SAPId = '" + SAPID + "'";

            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                CategoryTextBox.Text = reader[1].ToString();
            }

            // SubCategory
            cmd.CommandText = "SELECT P.SubCategory, C.SubCategoryName FROM PartsList P" +
            " INNER JOIN SubCategory C" +
            " ON P.SubCategory = C.SubCategoryID" +
            " WHERE SAPId = '" + SAPID + "'";

            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                SubCategoryTextBox.Text = reader[1].ToString();
            }

            // SSPCategory
            //cmd.CommandText = "SELECT P.SSPCategory, C.SSPCategoryName FROM PartsList P" +
            //" INNER JOIN SSPCategory C" +
            //" ON P.SSPCategory = C.SSPCategoryID" +
            //" WHERE SAPId = '" + SAPID + "'";

            //reader = cmd.ExecuteReader();
            //while (reader.Read())
            //{
            //    SSPCategoryTextBox.Text = reader[1].ToString();
            //}

            TecanDatabase.Close();
            reader.Dispose();
        }

        private void loadImage(string SAPID)
        {
            Byte[] imageData;
            try
            {
                openDB();
                SqlCeCommand cmd = TecanDatabase.CreateCommand();
                cmd.CommandText = "SELECT Document FROM PartImages WHERE SAPId = '" + SAPID + "'";
                imageData = (byte[])cmd.ExecuteScalar();
                if (imageData != null)
                {
                    Image newImage = byteArrayToImage(imageData);
                    newImage = ResizeImage(newImage, new Size(396, 224));
                    partImagePictureBox.Image = newImage;
                }
                else
                {
                    // If no image available
                    System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                    Stream myStream = myAssembly.GetManifestResourceStream("Tecan_Parts.noimage.bmp");
                    Bitmap image = new Bitmap(myStream);
                    System.Drawing.Image newImage = image;
                    newImage = ResizeImage(newImage, new Size(396, 224));
                    partImagePictureBox.Image = newImage;
                }
                    
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                TecanDatabase.Close();
            }
        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }

        public static Image ResizeImage(Image image, Size size, bool preserveAspectRatio = true)
        {
            int newWidth;
            int newHeight;
            if (preserveAspectRatio)
            {
                int originalWidth = image.Width;
                int originalHeight = image.Height;
                float percentWidth = (float)size.Width / (float)originalWidth;
                float percentHeight = (float)size.Height / (float)originalHeight;
                float percent = percentHeight < percentWidth ? percentHeight : percentWidth;
                newWidth = (int)(originalWidth * percent);
                newHeight = (int)(originalHeight * percent);
            }
            else
            {
                newWidth = size.Width;
                newHeight = size.Height;
            }
            Image newImage = new Bitmap(newWidth, newHeight);
            using (Graphics graphicsHandle = Graphics.FromImage(newImage))
            {
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight);
            }
            return newImage;
        }

        private void suppumentalDocsListBox_DoubleClick(object sender, EventArgs e)
        {
            String selectedDocName = suppumentalDocsListBox.GetItemText(suppumentalDocsListBox.SelectedItem);

            // Get the file contents from the database
            SqlCeConnection TecanSuppDocsDatabase = null;

            TecanSuppDocsDatabase = new SqlCeConnection();
            // String dataPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            TecanSuppDocsDatabase.ConnectionString = "Data Source=|DataDirectory|\\TecanSuppDocs.sdf;Max Database Size=4000;Max Buffer Size=1024;Persist Security Info=False";
            TecanSuppDocsDatabase.Open();
            SqlCeCommand cmd = TecanSuppDocsDatabase.CreateCommand();

            cmd.CommandText = "SELECT Document FROM SuppumentalDocs WHERE FileName = '" + selectedDocName + "'";
            SqlCeDataReader reader = cmd.ExecuteReader();
            reader = cmd.ExecuteReader();
            Byte[] documentData = new Byte[0];
            while (reader.Read())
            {
                documentData = (byte[])reader[0];
            }
            reader.Dispose();
            TecanSuppDocsDatabase.Close();

            // Create the new file in temp directory
            String tempFilePath = @AppDomain.CurrentDomain.BaseDirectory.ToString() + "temp";
            System.IO.Directory.CreateDirectory(tempFilePath);

            // If temp directory current contains any files, delete them
            System.IO.DirectoryInfo tempFiles = new DirectoryInfo(tempFilePath);

            foreach (FileInfo file in tempFiles.GetFiles())
            {
                file.Delete();
            }

            String fullFilePathName = @tempFilePath + "\\" + selectedDocName;
            System.IO.FileStream fs = System.IO.File.Create(fullFilePathName);
            fs.Close();

            // Write file contents into file
            BinaryWriter Writer = null;

            try
            {
                // Create a new stream to write to the file
                Writer = new BinaryWriter(File.OpenWrite(fullFilePathName));

                // Writer raw data                
                Writer.Write(documentData);
                Writer.Flush();
                Writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Process.Start(fullFilePathName);
        }

        private void loadRequiredParts(String SAPID)
        {
            RequiredListView.Items.Clear();
            
            openDB();
            SqlCeCommand cmd = TecanDatabase.CreateCommand();

            cmd.CommandText = "SELECT R.RequiredSAPId, P.Description FROM RequiredParts R" +
            " INNER JOIN PartsList P " +
            " ON R.RequiredSAPId = P.SAPId" +
            " WHERE R.SAPId = '" + SAPID + "'" +
            " ORDER BY RequiredSAPId";
            try
            {
                SqlCeDataReader reader = cmd.ExecuteReader();

                int partCount = 0;
                while (reader.Read())
                {
                    RequiredListView.Items.Add(reader[0].ToString());
                    RequiredListView.Items[partCount].SubItems.Add(reader[1].ToString());
                    partCount++;
                }
                reader.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            TecanDatabase.Close();

        }

        private void loadCompatibility(String SAPID)
        {
            compatibilityListBox.Items.Clear();
            String[] compatibilities = new String[20];

            openDB();
            SqlCeCommand cmd = TecanDatabase.CreateCommand();
            SqlCeDataReader reader;
            cmd.CommandText = "SELECT Compatibility FROM PartsList WHERE SAPId = '" + SAPID + "'";
            try
            {
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    compatibilities = reader[0].ToString().Split(',');
                }
                reader.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            for (int i = 0; i < compatibilities.Length; i++)
            {
                cmd.CommandText = "SELECT CompatibilityName FROM Compatibility WHERE CompatibilityID = '" + compatibilities[i] + "'";
                try
                {
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        compatibilityListBox.Items.Add(reader[0]);
                    }
                    reader.Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            TecanDatabase.Close();
            
        }

        private void loadPartsListLookupComboBox()
        {
            openDB();
            SqlCeCommand cmd = TecanDatabase.CreateCommand();
            SqlCeDataReader reader;
            cmd.CommandText = "SELECT SAPId, Description FROM PartsList";
            try
            {
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    PartListLookupComboBox.Items.Add(reader[0].ToString() + " - " + reader[1].ToString());
                }
                reader.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PartListLookupComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selectedListBoxItem;
            String selectedSAPID;
            Int32 endOfSAP;

            selectedListBoxItem = PartListLookupComboBox.GetItemText(PartListLookupComboBox.SelectedItem);
            endOfSAP = selectedListBoxItem.IndexOf(" -");
            selectedSAPID = selectedListBoxItem.Substring(0, endOfSAP);
            LoadParts(selectedSAPID);
        }

        //private void RequiredListView_Click(object sender, EventArgs e)
        //{
        //    String RequiredSAPID = RequiredListView.SelectedItems[0].Text;
        //    AlternativesListView.Items.Clear();
            
        //    String[] AlternateSAP = new String[10];
        //    String[] AlternateDescription = new String[10];

        //    openDB();
        //    SqlCeCommand cmd = TecanDatabase.CreateCommand();

        //    cmd.CommandText = "SELECT Alternatives FROM RequiredParts WHERE SAPId = '" + sAPIdTextBox.Text + "' AND RequiredSAPId = '" + RequiredSAPID + "'";
        //    try
        //    {
        //        SqlCeDataReader reader = cmd.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            AlternateSAP = reader[0].ToString().Split(',');
        //        }
        //        reader.Dispose();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }

        //    for (int i = 0; i < AlternateSAP.Length; i++)
        //    {
        //        cmd.CommandText = "SELECT Description FROM PartsList WHERE SAPId = '" + AlternateSAP[i] + "'";
        //        try
        //        {
        //            SqlCeDataReader reader = cmd.ExecuteReader();

        //            while (reader.Read())
        //            {
        //                AlternateDescription[i] = reader[0].ToString();
        //            }
        //            reader.Dispose();

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }
        //    }
        //    TecanDatabase.Close();
        //    for (int i = 0; i < AlternateSAP.Length; i++)
        //    {
        //        AlternativesListView.Items.Add(AlternateSAP[i]);
        //        AlternativesListView.Items[i].SubItems.Add(AlternateDescription[i]);

        //    }

        //}

        private void openDB()
        {
            TecanDatabase = new SqlCeConnection();
            String dataPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            TecanDatabase.ConnectionString = "Data Source=|DataDirectory|\\TecanCustomerPartsList.sdf;Max Database Size=4000;Max Buffer Size=1024;Persist Security Info=False";
            TecanDatabase.Open();
        }

        //private void standarPriceTextBox_TextChanged(object sender, EventArgs e)
        //{
        //    if (standarPriceTextBox.Focused != true)
        //    {
        //        if (standarPriceTextBox.Text.IndexOf("$") == -1)
        //        {
        //            standarPriceTextBox.Text = getFormatedDollarValue(standarPriceTextBox.Text);
        //        }
        //    }
        //}

        private void iLPTextBox_TextChanged(object sender, EventArgs e)
        {
            if (iLPTextBox.Focused != true)
            {
                if (iLPTextBox.Text.IndexOf("$") == -1)
                {
                    iLPTextBox.Text = getFormatedDollarValue(iLPTextBox.Text);
                }
            }
        }

        //private void manufacturingCostTextBox_TextChanged(object sender, EventArgs e)
        //{
        //    if (manufacturingCostTextBox.Focused != true)
        //    {
        //        if (manufacturingCostTextBox.Text.IndexOf("$") == -1)
        //        {
        //            manufacturingCostTextBox.Text = getFormatedDollarValue(manufacturingCostTextBox.Text);
        //        }
        //    }            
        //}

        //private void aSPTextBox_TextChanged(object sender, EventArgs e)
        //{
        //    if (aSPTextBox.Focused != true)
        //    {
        //        if (aSPTextBox.Text.IndexOf("$") == -1)
        //        {
        //            aSPTextBox.Text = getFormatedDollarValue(aSPTextBox.Text);
        //        }
        //    }
        //}
        
        //private void plPriceTextBox_TextChanged(object sender, EventArgs e)
        //{
        //    if (plPriceTextBox.Focused != true)
        //    {
        //        if (plPriceTextBox.Text.IndexOf("$") == -1)
        //        {
        //            plPriceTextBox.Text = getFormatedDollarValue(plPriceTextBox.Text);
        //        }
        //    }
        //}
        
        private String getFormatedDollarValue(String dollarString)
        {
            String dollardValue = dollarString;
            Decimal d = Decimal.Parse(dollardValue, NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, new CultureInfo("en-US"));
            dollardValue = String.Format("{0:C2}", d);
            return dollardValue;
        }

        private int GetDecimals(decimal d, int i = 0)
        {
            decimal multiplied = (decimal)((double)d * Math.Pow(10, i));
            if (Math.Round(multiplied) == multiplied)
                return i;
            return GetDecimals(d, i + 1);
        }

        //private void PartsListDetailDisplay_Load(object sender, EventArgs e)
        //{
        //    // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.SubCategory' table. You can move, or remove it, as needed.
        //    this.subCategoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.SubCategory);
        //    // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.Instrument' table. You can move, or remove it, as needed.
        //    this.instrumentTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Instrument);
        //    // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.Category' table. You can move, or remove it, as needed.
        //    this.categoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Category);
        //    // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.PartsList' table. You can move, or remove it, as needed.
        //    this.partsListTableAdapter.Fill(this.tecanCustomerPartsListDataSet.PartsList);

        //}

    }
}
