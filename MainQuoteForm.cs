using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.Net.Mail;
using System.Diagnostics;

namespace Tecan_Parts
{
    public partial class MainQuoteForm : Form
    {
        SqlCeConnection TecanDatabase = null;

        Boolean searchPreformed = true;
        //Boolean salesTypeChanged = false;
        Boolean instrumentChanged = false;
        Boolean categoryChanged = false;
        // Boolean subCategoryChanged = false;
        // Boolean formatOnly = false;

        PartsListDetailDisplay DetailsForm;
        Profile profile = new Profile();
        // private System.Drawing.Font printFont;
        // private StreamReader streamToPrint;

        public MainQuoteForm()
        {
            InitializeComponent();
        }

        public void MainQuoteForm_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.SubCategory' table. You can move, or remove it, as needed.
            this.subCategoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.SubCategory);
            // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.Category' table. You can move, or remove it, as needed.
            this.categoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Category);
            // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.Instrument' table. You can move, or remove it, as needed.
            this.instrumentTableAdapter.Fill(this.tecanCustomerPartsListDataSet.Instrument);
            // TODO: This line of code loads data into the 'tecanCustomerPartsListDataSet.PartsList' table. You can move, or remove it, as needed.
            this.partsListTableAdapter.Fill(this.tecanCustomerPartsListDataSet.PartsList);
            // TODO: This line of code loads data into the 'TecanCustomerPartsListDataSet11.SubCategory' table. You can move, or remove it, as needed.
            this.subCategoryTableAdapter.Fill(this.tecanCustomerPartsListDataSet.SubCategory);
            // TODO: This line of code loads data into the 'TecanCustomerPartsListDataSet11.Category' table. You can move, or remove it, as needed.
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.PartNumberClearButton, "Clear Part Number Search");
            ToolTip1.SetToolTip(this.DescriptionClearButton, "Clear Description Search");
            ToolTip1.SetToolTip(this.ClearFiltersButton, "Reset All Category Filters");

            // Check for salesman profile, if no file get salesman information
            // Check if Quote Database is empty
            if (partsListBindingSource.Count != 0)
            {
                loadFilterComboBoxes();
                int newGridHeight;
                newGridHeight = Screen.PrimaryScreen.Bounds.Height - this.menuStrip1.Height;
                // this.partsListDataGridView.Height = newGridHeight - 500;
                this.partsListDataGridView.Height = 425;
                setPartDetailTextBox();
                OptionsDataGridView.AllowDrop = true;
                QuoteTabControl.SelectedTab = QuoteSettingTabPage;
            }
        }

        // Called from Form Shown Event!
        // Reads or Creates users profile xml file.
        private void getProfileAndDatabase(object sender, EventArgs e)
        {
            String profileFile = @"c:\TecanFiles\" + "TecanConfig.cfg";
            // Normal Condition, just load profile and run
            if (File.Exists(profileFile) && partsListBindingSource.Count > 1)
            {
                getUsersProfile();
            }
            // Checks DB and copies / loads if required
            else if (partsListBindingSource.Count == 1)
            {
                showUserProfileForm(true);
            }
            else if (!File.Exists(profileFile))
            {
                showUserProfileForm(false);
            }

            // Checks DB and copies / loads if required
            //if (partsListBindingSource.Count == 1)
            //{
            //    String distributionFolder = "";
            //    distributionFolder = profile.DistributionFolder;
            //    Boolean fileFound = false;
            //    fileFound = copyDatabaseToWorkingFolder(distributionFolder);
            //    while(!fileFound)
            //    {
            //        MessageBox.Show("The Distribution Folder you selected in your profile does not contain the Parts List Database!\n\nPlease select a new folder");
            //        showUserProfileForm(true);
            //        distributionFolder = profile.DistributionFolder;
            //        fileFound = copyDatabaseToWorkingFolder(distributionFolder);
            //    }
            //}
            //// If = 1 then no DB, requires intilization
            //if (partsListBindingSource.Count == 1)
            //{

            //    if (MessageBox.Show("The Tecan Parts List must be intilized!\r\n\r\nDo you want to perform intialization now?", "Initial Installation", MessageBoxButtons.YesNo) == DialogResult.No)
            //    {
            //        this.Close();
            //    }
            //    else
            //    {
            //        if (!File.Exists(profileFile))
            //        {
            //            ProfileForm profileForm = new ProfileForm(true);
            //            profileForm.SetForm1Instance(this);
            //            profileForm.Show();
            //            Application.OpenForms["ProfileForm"].BringToFront();
            //        }
            //        else
            //        {
            //            // Profile exists, get distrubtion folder and then database
            //            String distributionFolder;
            //            getUsersProfile();
            //            distributionFolder = profile.DistributionFolder;

            //            if (distributionFolder == null)
            //            {
            //                ProfileForm profileForm = new ProfileForm(true);
            //                profileForm.SetForm1Instance(this);
            //                profileForm.Show();
            //                Application.OpenForms["ProfileForm"].BringToFront();
            //                MessageBox.Show("There's a problem with your profile settings.  Please re-enter and save your information!");
            //            }

            //            Boolean fileFound;
            //            fileFound = copyDatabaseToWorkingFolder(distributionFolder);
            //            if (!fileFound)
            //            {
            //                MessageBox.Show("The Distribution Folder you selected in your profile does not contain the Parts List Database!\n\nPlease select a new folder");
            //                ProfileForm profileForm = new ProfileForm(true);
            //                profileForm.SetForm1Instance(this);
            //                profileForm.Show();
            //                Application.OpenForms["ProfileForm"].BringToFront();
            //            }
            //            else
            //            {
            //                MainQuoteForm_Load(sender, e);
            //            }
            //        }
            //    }
            //}
            //else
            //{
            //    if (!File.Exists(profileFile))
            //    {
            //        ProfileForm profileForm = new ProfileForm(false);
            //        profileForm.SetForm1Instance(this);
            //        profileForm.Show();
            //        Application.OpenForms["ProfileForm"].BringToFront();
            //    }
            //    else
            //    {
            //        getUsersProfile();
            //    }
            //}
        }

        public void showUserProfileForm(Boolean NeedsDB)
        {
            ProfileForm profileForm = new ProfileForm(NeedsDB);
            profileForm.SetForm1Instance(this);
            profileForm.Show();
            Application.OpenForms["ProfileForm"].BringToFront();
        }

        // If blank database or new database available copy new database to working folder
        // todo See if Supplemental database is needed
        public Boolean copyDatabaseToWorkingFolder(String sourcePath)
        {
            String quoteSourceFile = "";
            String supplementSourceFile;
            
            // Where new files will go
            String quoteTargetFile = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "TecanCustomerPartsList.sdf");
            String supplementTargetFile = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "TecanSuppDocs.sdf");

            // Where the new files will come from
            try
            {
                quoteSourceFile = System.IO.Path.Combine(sourcePath, "TecanCustomerPartsList.sdf");
            }
                catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            supplementSourceFile = System.IO.Path.Combine(sourcePath, "TecanSuppDocs.sdf");

            // Verify the files exisit before copy
            if (!File.Exists(quoteSourceFile)) { 
                return false; 
            }

            getUsersProfile();
            FileInfo fi = new FileInfo(quoteSourceFile);
            profile.DatabaseCreationDate = fi.CreationTime;
            saveUsersProfile();
            System.IO.File.Copy(quoteSourceFile, quoteTargetFile, true);
            System.IO.File.Copy(supplementSourceFile, supplementTargetFile, true);
            return true;
        }

        // Initial Lookup Table lists, used for filtering displayed Parts
        // Lists are all loaded from lookup table items with all items
        // Look at updateCategoryComboBox() and updateSubCategoryComboBox() if only avaible items are required
        private void loadFilterComboBoxes()
        {
            // Sales Type
            //SalesTypeComboBox.DataSource = this.SalesTypeBindingSource;
            //SalesTypeComboBox.DisplayMember = "SalesTypeName";
            //SalesTypeComboBox.ValueMember = "SalesTypeID";
            //if (this.SalesTypeBindingSource.Count > 0) SalesTypeComboBox.SelectedIndex = 0;

            // Instrument
            InstrumentComboBox.DataSource = this.InstrumentBindingSource;
            InstrumentComboBox.DisplayMember = "InstrumentName";
            InstrumentComboBox.ValueMember = "InstrumentID";
            if (this.InstrumentBindingSource.Count > 0) InstrumentComboBox.SelectedIndex = 0;

            // Category
            CategoryComboBox.DataSource = this.CategoryBindingSource;
            CategoryComboBox.DisplayMember = "CategoryName";
            CategoryComboBox.ValueMember = "CategoryID";
            if (this.CategoryBindingSource.Count > 0) CategoryComboBox.SelectedIndex = 0;

            // Sub Categories
            SubCategoryComboBox.DataSource = this.SubCategoryBindingSource;
            SubCategoryComboBox.DisplayMember = "SubCategoryName";
            SubCategoryComboBox.ValueMember = "SubCategoryID";
            if (this.SubCategoryBindingSource.Count > 0) SubCategoryComboBox.SelectedIndex = 0;
        }

        private void setPartDetailTextBox()
        {
            String CatchEmpty = this.partsListDataGridView.CurrentCell.Value.ToString();
            if (CatchEmpty != "")
            {
                System.Data.DataRowView SelectedRowView;
                TecanCustomerPartsListDataSet.PartsListRow SelectedRow;
                SelectedRowView = (System.Data.DataRowView)partsListBindingSource.Current;
                SelectedRow = (TecanCustomerPartsListDataSet.PartsListRow)SelectedRowView.Row;
                PartDetailTextBox.Text = SelectedRow.DetailDescription;
                itemPriceTextBox.Text = String.Format("{0:C2}", SelectedRow.ILP);
                NotesTextBox.Text = SelectedRow.NotesFromFile;
                loadImage(SelectedRow.SAPId);
            }
        }

        // Ensures the selected item is displayed in the detail and image info
        private void partsListDataGridView_MouseMove(object sender, MouseEventArgs e)
        {
            Int32 selectedRowCount = partsListDataGridView.Rows.GetRowCount(DataGridViewElementStates.Selected);
            DataGridView.HitTestInfo info = partsListDataGridView.HitTest(e.X, e.Y);
            
            // set the currentcell property manually.   
            if (info.RowIndex >= 0 && selectedRowCount == 1)
            {
                try
                {
                    this.partsListDataGridView.CurrentCell = this.partsListDataGridView[info.ColumnIndex, info.RowIndex];
                }
                catch { }
                setPartDetailTextBox();
            }
        }

        // Event Trigers Drag Drop (originator)
        private void partsListDataGridView_MouseDown(object sender, MouseEventArgs e)
        {
            Boolean CatchError = false;
            if (e.Button == MouseButtons.Left)
            {
                if (QuoteTabControl.SelectedTab == QuoteSettingTabPage)
                {
                    QuoteTabControl.SelectedTab = OptionTabPage;
                }
                partsListDataGridView.DoDragDrop(partsListDataGridView.SelectedRows, DragDropEffects.Move);
            }
            else
            {
                System.Data.DataRowView SelectedRowView;
                TecanCustomerPartsListDataSet.PartsListRow SelectedRow;

                SelectedRowView = (System.Data.DataRowView)partsListBindingSource.Current;
                SelectedRow = (TecanCustomerPartsListDataSet.PartsListRow)SelectedRowView.Row;

                if (DetailsForm == null || DetailsForm.IsDisposed) DetailsForm = new PartsListDetailDisplay();
                try
                {
                    DetailsForm.LoadParts(SelectedRow.SAPId);
                }
                catch (Exception ex)
                {
                    CatchError = true;
                }
                if (!CatchError)
                {
                    DetailsForm.SetForm1Instance(this);
                    DetailsForm.Show();
                    Application.OpenForms["PartsListDetailDisplay"].BringToFront();
                }
            }
        }


        private void SumItems(DataGridView myDataGridView)
        {
            Int32 rowCount = myDataGridView.Rows.GetRowCount(DataGridViewElementStates.Displayed);
            Int32 rowIndex;
            Decimal itemPrice = 0;

            for (int s = 0; s < rowCount; s++)
            {
                rowIndex = myDataGridView.Rows[s].Index;
                DataGridViewRow srow = myDataGridView.Rows[rowIndex];
                itemPrice = itemPrice + (Decimal)srow.Cells[4].Value;
            }

            switch (myDataGridView.Name)
            {
                case "QuoteDataGridViewM":
                    // QuoteItemsPriceTextBox.Text = String.Format("{0:C2}", itemPrice); // getFormatedDollarValue(itemPrice.ToString());
                    break;

                case "OptionsDataGridView":
                    OptionsItemsPriceTextBox.Text = String.Format("{0:C2}", itemPrice); //getFormatedDollarValue(itemPrice.ToString());
                    break;
            }
        }

        private void processCellValueChange(DataGridView myDataGridView, DataGridViewCellEventArgs e)
        {
            Decimal itemPrice;
            Int32 itemQty;
            // Decimal itemDiscount;
            // Decimal discountPercentage;
            Decimal extendedPrice;
            String itemPriceCheck;
            // String discountCheck;

            itemPriceCheck = myDataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            // discountCheck = myDataGridView.Rows[e.RowIndex].Cells[4].Value.ToString();

            //if (itemPriceCheck.IndexOf("$") == -1)
            //{
            //    formatOnly = false;
            //}

            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {
                if (itemPriceCheck.IndexOf("$") != -1) itemPriceCheck = itemPriceCheck.Substring(1, itemPriceCheck.Length - 1);
                itemPrice = Convert.ToDecimal(itemPriceCheck);

                // MessageBox.Show(myDataGridView.Rows[e.RowIndex].Cells[3].Value.ToString());
                itemQty = Convert.ToInt32(myDataGridView.Rows[e.RowIndex].Cells[3].Value);
                extendedPrice = itemPrice * itemQty;
                myDataGridView.Rows[e.RowIndex].Cells[4].Value = extendedPrice;
                SumItems(myDataGridView);
            }

            if (itemPriceCheck.IndexOf("$") == -1)
            {
                // formatOnly = true;
                itemPrice = Convert.ToDecimal(itemPriceCheck);
                OptionsDataGridView.Rows[e.RowIndex].Cells[2].Value = String.Format("{0:C2}", itemPrice);
            }

        }

        private void OptionsDataGridView_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(DataGridViewSelectedRowCollection)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        // The drop into the desired object
        private void OptionsDataGridView_DragDrop(object sender, DragEventArgs e)
        {
            String itemSAPID;
            String itemDescription;
            Decimal itemPrice;

            DataGridViewSelectedRowCollection rows = (DataGridViewSelectedRowCollection)e.Data.GetData(typeof(DataGridViewSelectedRowCollection));
            foreach (DataGridViewRow row in rows)
            {
                itemSAPID = row.Cells[0].Value.ToString();
                itemDescription = row.Cells[1].Value.ToString();
                itemPrice = getPartPrice(itemSAPID);
                OptionsDataGridView.Rows.Add(itemSAPID, itemDescription, itemPrice, 1, itemPrice);
            }
            SumItems(OptionsDataGridView);
        }

        // Update Extended Price and Totals when QTY and/or Discount Change
        private void OptionsDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                processCellValueChange(OptionsDataGridView, e);
            }

        }

        public Decimal getPartPrice(String SAPID)
        {
            Decimal itemPrice = 0;
            openDB();
            SqlCeCommand cmd = TecanDatabase.CreateCommand();
            cmd.CommandText = "SELECT ILP FROM PartsList WHERE SAPId = '" + SAPID + "'";
            try
            {
                itemPrice = (Decimal)cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            return itemPrice;
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
                    System.Drawing.Image newImage = byteArrayToImage(imageData);
                    newImage = ResizeImage(newImage, new Size(123, 135));
                    partImagePictureBox.Image = newImage;
                }
                else
                {
                    partImagePictureBox.Image = null;
                }

            }
            finally
            {
                TecanDatabase.Close();
            }
        }

        public System.Drawing.Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream ms = new MemoryStream(byteArrayIn);
            System.Drawing.Image returnImage = null;
            try
            {
                returnImage = System.Drawing.Image.FromStream(ms);
            }
            catch { }
            return returnImage;
        }

        private void openDB()
        {
            TecanDatabase = new SqlCeConnection();
            String dataPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            TecanDatabase.ConnectionString = "Data Source=|DataDirectory|\\TecanCustomerPartsList.sdf;Max Database Size=4000;Max Buffer Size=1024;Persist Security Info=False";
            TecanDatabase.Open();
        }


        public static System.Drawing.Image ResizeImage(System.Drawing.Image image, Size size, bool preserveAspectRatio = true)
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
            System.Drawing.Image newImage = new Bitmap(newWidth, newHeight);
            using (Graphics graphicsHandle = Graphics.FromImage(newImage))
            {
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight);
            }
            return newImage;
        }

        private void doSearch()
        {
            if (!searchPreformed)
            {
                searchPreformed = true;
                String PartSearchValue;
                String DescriptionSearchValue;

                short InstrumentSearchValue = (short)Convert.ToInt16(InstrumentComboBox.SelectedValue);
                short CategorySearchValue = (short)Convert.ToInt16(CategoryComboBox.SelectedValue);
                short SubCategorySearchValue = (short)Convert.ToInt16(SubCategoryComboBox.SelectedValue);
                // byte SalesTypeSearchValue = (byte)Convert.ToInt16(SalesTypeComboBox.SelectedValue);
                // byte SalesTypeSearchValue = 0;


                // Set Text Search values
                if (PartNumberSearchTextBox.Text != "")
                {
                    PartSearchValue = PartNumberSearchTextBox.Text + "%";
                }
                else
                {
                    PartSearchValue = "%%";
                }

                if (DescriptionSearchTextBox.Text != "")
                {
                    DescriptionSearchValue = "%" + DescriptionSearchTextBox.Text + "%";
                }
                else
                {
                    DescriptionSearchValue = "%%";
                }

                // Test Filter Values values
                // No Filters Set
                if ((InstrumentSearchValue == 0) && (CategorySearchValue == 0) && (SubCategorySearchValue == 0))
                {
                    partsListTableAdapter.FillByLIKE(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue);
                }
                // Instrument Only
                else if ((InstrumentSearchValue != 0) && (CategorySearchValue == 0) && (SubCategorySearchValue == 0))
                {
                    partsListTableAdapter.FillByLIKEANDInstrument(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue);
                }
                // Category Only
                else if ((InstrumentSearchValue == 0) && (CategorySearchValue != 0) && (SubCategorySearchValue == 0))
                {
                    partsListTableAdapter.FillByLIKEANDCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, CategorySearchValue);
                }
                // Sub Category Only
                else if ((InstrumentSearchValue == 0) && (CategorySearchValue == 0) && (SubCategorySearchValue != 0))
                {
                    partsListTableAdapter.FillByLIKEANDSubCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, SubCategorySearchValue);
                }
                // Sales Type Only
                //else if ((InstrumentSearchValue == 0) && (CategorySearchValue == 0) && (SubCategorySearchValue == 0) && (SalesTypeSearchValue != 0))
                //{
                //    partsListTableAdapter.FillByLIKEANDSalesType(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, SalesTypeSearchValue);
                //}
                // Instrunment Combinations
                // Instrument AND Category
                else if ((InstrumentSearchValue != 0) && (CategorySearchValue != 0) && (SubCategorySearchValue == 0))
                {
                    partsListTableAdapter.FillByLIKEANDInstrumentANDCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, CategorySearchValue);
                }
                // Instrument AND SubCategory
                else if ((InstrumentSearchValue != 0) && (CategorySearchValue == 0) && (SubCategorySearchValue != 0))
                {
                    partsListTableAdapter.FillByLIKEANDInstrumentANDSubCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, SubCategorySearchValue);
                }
                // Instrument AND Sales Type
                //else if ((InstrumentSearchValue != 0) && (CategorySearchValue == 0) && (SubCategorySearchValue == 0) && (SalesTypeSearchValue != 0))
                //{
                //    partsListTableAdapter.FillByLIKEANDInstrumentANDSalesType(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, SalesTypeSearchValue);
                //}
                // Instrument AND Category AND Sub Category
                else if ((InstrumentSearchValue != 0) && (CategorySearchValue != 0) && (SubCategorySearchValue != 0))
                {
                    partsListTableAdapter.FillByLIKEANDInstrumentANDCategoryANDSubCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, CategorySearchValue, SubCategorySearchValue);
                }
                // Instrument AND Category AND Sales Type
                //else if ((InstrumentSearchValue != 0) && (CategorySearchValue != 0) && (SubCategorySearchValue == 0) && (SalesTypeSearchValue != 0))
                //{
                //    partsListTableAdapter.FillByLIKEANDInstrumentANDCategoryANDSalesType(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, CategorySearchValue, SalesTypeSearchValue);
                //}
                // Instrument AND SubCategory AND Sales Type
                //else if ((InstrumentSearchValue != 0) && (CategorySearchValue == 0) && (SubCategorySearchValue != 0) && (SalesTypeSearchValue != 0))
                //{
                //    partsListTableAdapter.FillByLIKEANDInstrumentANDSubCategoryANDSalesType(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, SubCategorySearchValue, SalesTypeSearchValue);
                //}
                // Instrument AND Category AND SubCategory AND Sales Type
                //else if ((InstrumentSearchValue != 0) && (CategorySearchValue != 0) && (SubCategorySearchValue != 0) && (SalesTypeSearchValue != 0))
                //{
                //    partsListTableAdapter.FillByLIKEANDInstrumentANDCategoryANDSubCategoryANDSalesType(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, InstrumentSearchValue, CategorySearchValue, SubCategorySearchValue, SalesTypeSearchValue);
                //}
                // Category Combinations
                // Category AND SubCategory
                else if ((InstrumentSearchValue == 0) && (CategorySearchValue != 0) && (SubCategorySearchValue != 0))
                {
                    partsListTableAdapter.FillByLIKEANDCategoryANDSubCategory(this.tecanCustomerPartsListDataSet.PartsList, PartSearchValue, DescriptionSearchValue, CategorySearchValue, SubCategorySearchValue);
                }

                if(instrumentChanged)
                {
                    // updateSalesTypeComboBox();
                    updateCategoryComboBox();
                    updateSubCategoryComboBox();
                }
                if(categoryChanged)
                {
                    // updateSalesTypeComboBox();
                    // updateInstrumentComboBox();
                    updateSubCategoryComboBox();
                }

            }
        }

        // Called after doSearch() if Category or SubCategory has changed
        // Only include Instruments that are avaiable for selected Dataset
        private void updateInstrumentComboBox()
        {
            // Instruments
            //salesTypeChanged = false;
            categoryChanged = false;
            // subCategoryChanged = false;
            String CurrentInstrumentSearchValue = InstrumentComboBox.SelectedValue.ToString();
            // short SalesTypeSearchValue = (short)Convert.ToInt16(SalesTypeComboBox.SelectedValue);
            short SalesTypeSearchValue = 0;

            short CategorySearchValue = (short)Convert.ToInt16(CategoryComboBox.SelectedValue);
            short SubCategorySearchValue = (short)Convert.ToInt16(SubCategoryComboBox.SelectedValue);

            if (SalesTypeSearchValue != 0 || CategorySearchValue != 0 || SubCategorySearchValue != 0)
            {
                // Set up new Available Category list
                ArrayList theAvailableInstruments = new ArrayList();
                theAvailableInstruments.Add(new LookupTableDefinitions.AvailableInstruments("Any", "0"));

                foreach (DataRow dr in this.tecanCustomerPartsListDataSet.Tables["Instrument"].Rows)
                {
                    foreach (DataRow pr in this.tecanCustomerPartsListDataSet.Tables["PartsList"].Rows)
                    {
                        if (pr["Instrument"].ToString() == dr["InstrumentID"].ToString())
                        {
                            theAvailableInstruments.Add(new LookupTableDefinitions.AvailableInstruments(dr["InstrumentName"].ToString(), dr["InstrumentID"].ToString()));
                            break;
                        }
                    }
                }
                // InstrumentListComboBox
                InstrumentComboBox.DataSource = theAvailableInstruments;
                InstrumentComboBox.DisplayMember = "Name";
                InstrumentComboBox.ValueMember = "ID";
            }
            else
            {
                // InstrumentListComboBox
                InstrumentComboBox.DataSource = this.InstrumentBindingSource;
                InstrumentComboBox.DisplayMember = "InstrumentName";
                InstrumentComboBox.ValueMember = "InstrumentID";
            }

            // Set the perviously selected item
            if (CurrentInstrumentSearchValue != "0")
            {
                searchPreformed = true;
                int currentItem = 0;

                try
                {
                    foreach (LookupTableDefinitions.AvailableInstruments instruments in InstrumentComboBox.Items)
                    {
                        if (instruments.ID == CurrentInstrumentSearchValue)
                        {
                            InstrumentComboBox.SelectedIndex = currentItem;
                        }
                        currentItem++;
                    }
                }
                catch (Exception)
                { }

            }

        }

        // Called after doSearch() if Instrument has changed
        // Only include Categories that are avaiable for selected Instrument
        private void updateCategoryComboBox()
        {
            // Categories
            String CurrentCategorySearchValue = CategoryComboBox.SelectedValue.ToString();
            //salesTypeChanged = false;
            instrumentChanged = false;
            // short SalesTypeSearchValue = (short)Convert.ToInt16(SalesTypeComboBox.SelectedValue);
            short SalesTypeSearchValue = 0;
            short InstrumentSearchValue = (short)Convert.ToInt16(InstrumentComboBox.SelectedValue);
            short SubCategorySearchValue = (short)Convert.ToInt16(SubCategoryComboBox.SelectedValue);

            if (SalesTypeSearchValue != 0 || InstrumentSearchValue != 0 || SubCategorySearchValue != 0)
            {
                // Set up new Available Category list
                ArrayList theAvailableCategories = new ArrayList();
                theAvailableCategories.Add(new LookupTableDefinitions.AvailableCategories("Any", "0"));

                foreach (DataRow dr in this.tecanCustomerPartsListDataSet.Tables["Category"].Rows)
                {
                    foreach (DataRow pr in this.tecanCustomerPartsListDataSet.Tables["PartsList"].Rows)
                    {
                        if (pr["Category"].ToString() == dr["CategoryID"].ToString())
                        {
                            theAvailableCategories.Add(new LookupTableDefinitions.AvailableCategories(dr["CategoryName"].ToString(), dr["CategoryID"].ToString()));
                            break;
                        }
                    }
                }
                // CategoryListComboBox
                CategoryComboBox.DataSource = theAvailableCategories;
                CategoryComboBox.DisplayMember = "Name";
                CategoryComboBox.ValueMember = "ID";
            }
            else
            {
                // CategoryListComboBox
                CategoryComboBox.DataSource = this.CategoryBindingSource;
                CategoryComboBox.DisplayMember = "CategoryName";
                CategoryComboBox.ValueMember = "CategoryID";
            }

            // Set the perviously selected item
            if (CurrentCategorySearchValue != "0")
            {
                searchPreformed = true;
                int currentItem = 0;

                try
                {
                    foreach (LookupTableDefinitions.AvailableCategories category in CategoryComboBox.Items)
                    {
                        if (category.ID == CurrentCategorySearchValue)
                        {
                            CategoryComboBox.SelectedIndex = currentItem;
                        }
                        currentItem++;
                    }
                }
                catch (Exception)
                { }
            }

        }

        // Called after doSearch() if Category has changed
        // Only include Sub-Categories that are avaiable for selected Category
        private void updateSubCategoryComboBox()
        {
            // SubCategories
            if (SubCategoryComboBox.SelectedValue == null) return;
            String CurrentSubCategorySearchValue = SubCategoryComboBox.SelectedValue.ToString();
            categoryChanged = false;
            //salesTypeChanged = false;
            instrumentChanged = false;
            // short SalesTypeSearchValue = (short)Convert.ToInt16(SalesTypeComboBox.SelectedValue);
            short SalesTypeSearchValue = 0;
            short InstrumentSearchValue = (short)Convert.ToInt16(InstrumentComboBox.SelectedValue);
            short CategorySearchValue = (short)Convert.ToInt16(CategoryComboBox.SelectedValue);

            if (SalesTypeSearchValue != 0 || InstrumentSearchValue != 0 || CategorySearchValue != 0)
            {
                // Set up new Available SubCategory list
                ArrayList theAvailableSubCategories = new ArrayList();
                theAvailableSubCategories.Add(new LookupTableDefinitions.AvailableSubCategories("Any", "0"));

                foreach (DataRow dr in this.tecanCustomerPartsListDataSet.Tables["SubCategory"].Rows)
                {
                    foreach (DataRow pr in this.tecanCustomerPartsListDataSet.Tables["PartsList"].Rows)
                    {
                        if (pr["SubCategory"].ToString() == dr["SubCategoryID"].ToString())
                        {
                            theAvailableSubCategories.Add(new LookupTableDefinitions.AvailableSubCategories(dr["SubCategoryName"].ToString(), dr["SubCategoryID"].ToString()));
                            break;
                        }
                    }
                }
                // SubCategoryListComboBox
                SubCategoryComboBox.DataSource = theAvailableSubCategories;
                SubCategoryComboBox.DisplayMember = "Name";
                SubCategoryComboBox.ValueMember = "ID";
            }
            else
            {
                // SubCategoryListComboBox
                SubCategoryComboBox.DataSource = this.SubCategoryBindingSource;
                SubCategoryComboBox.DisplayMember = "SubCategoryName";
                SubCategoryComboBox.ValueMember = "SubCategoryID";
            }

            // Set the perviously selected item
            if (CurrentSubCategorySearchValue != "0")
            {
                searchPreformed = true;
                int currentItem = 0;

                try
                {
                    foreach (LookupTableDefinitions.AvailableSubCategories category in SubCategoryComboBox.Items)
                    {
                        if (category.ID == CurrentSubCategorySearchValue)
                        {
                            SubCategoryComboBox.SelectedIndex = currentItem;
                        }
                        currentItem++;
                    }
                }
                catch (Exception)
                { }
            }
        }

        private void PartNumberSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            searchPreformed = false;
            doSearch();
        }

        private void PartNumberClearButton_Click(object sender, EventArgs e)
        {
            PartNumberSearchTextBox.Text = "";
            searchPreformed = false;
            doSearch();
        }

        private void DescriptionClearButton_Click(object sender, EventArgs e)
        {
            DescriptionSearchTextBox.Text = "";
            searchPreformed = false;
            doSearch();
        }

        private void DescriptionSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            searchPreformed = false;
            doSearch();
        }

        private void InstrumentComboBox_Click(object sender, EventArgs e)
        {
            searchPreformed = false;
            instrumentChanged = true;
        }

        private void InstrumentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CategoryComboBox.SelectedIndex != -1)
            {
                CategoryComboBox.SelectedIndex = 0;
                SubCategoryComboBox.SelectedIndex = 0;
                doSearch();
            }
        }

        private void CategoryComboBox_Click(object sender, EventArgs e)
        {
            searchPreformed = false;
            categoryChanged = true;
        }

        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SubCategoryComboBox.SelectedIndex != 0 && SubCategoryComboBox.SelectedIndex != -1) SubCategoryComboBox.SelectedIndex = 0; 
            doSearch();
        }

        private void SubCategoryComboBox_Click(object sender, EventArgs e)
        {
            searchPreformed = false;
            // subCategoryChanged = true;
        }

        private void SubCategoryComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            doSearch();
        }

        private void ClearFiltersButton_Click(object sender, EventArgs e)
        {
            loadFilterComboBoxes();
            searchPreformed = false;
            doSearch();
        }

        private void accountListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ContactsForm contacts = new ContactsForm();
            // contacts.Show();
        }

        private void OptionsRemoveSelectedButton_Click(object sender, EventArgs e)
        {
            RemoveItems(OptionsDataGridView);
            SumItems(OptionsDataGridView);
        }

        private void RemoveItems(DataGridView myDataGridView)
        {
            Int32 selectedRowCount = myDataGridView.Rows.GetRowCount(DataGridViewElementStates.Selected);
            Int32 removedRowCount = 0;
            Int32 totalRowCount = myDataGridView.RowCount;
            if (selectedRowCount > 0)
            {
                if (selectedRowCount == totalRowCount)
                {
                    MessageBox.Show("All cells are selected", "Selected Cells");
                    myDataGridView.Rows.Clear();
                }
                else
                {
                    while (selectedRowCount > removedRowCount)
                    {
                        myDataGridView.Rows.RemoveAt(myDataGridView.SelectedCells[0].RowIndex);
                        removedRowCount++;
                    }
                }
            }
        }

        private void saveQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (QuoteTitleTextBox.Text == "")
            {
                QuoteTabControl.SelectedTab = QuoteSettingTabPage;
                QuoteTitleTextBox.Focus();
                MessageBox.Show("Please enter a Order Title before saving.");
                return;
            }
            String QuoteFileName = QuoteTitleTextBox.Text + " - " + QuoteDateTimePicker.Text.Replace('/','_') + ".tbq";
            if (File.Exists(@"c:\TecanFiles\" + QuoteFileName))
            {
                if (MessageBox.Show("The order file named " + QuoteFileName + " already exists!\r\n\r\nDo you want to overwrite this file?", "Replace File", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    QuoteTabControl.SelectedTab = QuoteSettingTabPage;
                    QuoteTitleTextBox.Focus();
                    MessageBox.Show("Please enter a new Order Title.");
                    return;
                }
            }
            updateActionStatus("Saving order " + QuoteFileName);
            Application.DoEvents();
            Quote quote = new Quote();
            quote.QuoteTitle = QuoteTitleTextBox.Text;
            quote.QuoteDate = QuoteDateTimePicker.Value;
            quote.QuoteEmailTo = ProfileTecanEmailTextBox.Text;
            // quote.Items = AddQuoteItems(QuoteDataGridViewM);
            quote.Items = AddQuoteItems(OptionsDataGridView);
            // quote.ThirdParty = AddQuoteItems(ThirdPartyDataGridView);
            // quote.SmartStart = AddQuoteItems(SmartStartDataGridView);

            // Sum all items in order and set order total
            decimal quoteTotal = 0;
            foreach (QuoteItems row in quote.Items)
            {
                quoteTotal = quoteTotal + (row.Price * row.Quantity);
            }
            quote.QuoteTotal = quoteTotal;

            System.Xml.Serialization.XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(Quote));
            // Create the new directory if does not exist
            String tecanFilesFilePath = @"c:\TecanFiles";
            System.IO.Directory.CreateDirectory(tecanFilesFilePath);

            System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\TecanFiles\" + QuoteFileName);
            writer.Serialize(file, quote);
            file.Close();
            actionStatusLabel.Text = "";
        }

        private ArrayList AddQuoteItems(DataGridView myDataGridView)
        {
            ArrayList quoteItems = new ArrayList();

            String SAPID;
            String Description;
            Decimal itemPrice;
            Int32 itemQty;
            //Boolean Note;
            //Boolean Image;
            //Decimal discountPercentage;
            String itemPriceCheck;
            //String discountCheck;
            Int16 rowCount = 0;

            foreach (DataGridViewRow row in myDataGridView.Rows)
            {
                SAPID = myDataGridView.Rows[rowCount].Cells[0].Value.ToString();
                Description = myDataGridView.Rows[rowCount].Cells[1].Value.ToString();

                itemPriceCheck = myDataGridView.Rows[rowCount].Cells[2].Value.ToString();
                // discountCheck = myDataGridView.Rows[rowCount].Cells[4].Value.ToString();

                if (itemPriceCheck.IndexOf("$") != -1) itemPriceCheck = itemPriceCheck.Substring(1, itemPriceCheck.Length - 1);
                itemPrice = Convert.ToDecimal(itemPriceCheck);

                //if (discountCheck.IndexOf("%") != -1) discountCheck = discountCheck.Substring(0, discountCheck.Length - 2);
                //discountPercentage = Convert.ToDecimal(discountCheck);

                itemQty = Convert.ToInt32(myDataGridView.Rows[rowCount].Cells[3].Value);

                //Note = Convert.ToBoolean(myDataGridView.Rows[rowCount].Cells[6].Value);
                //Image = Convert.ToBoolean(myDataGridView.Rows[rowCount].Cells[7].Value);

                QuoteItems newItem = new QuoteItems();
                newItem.SAPID = SAPID;
                newItem.Description = Description;
                newItem.Quantity = itemQty;
                newItem.Price = itemPrice;
                //newItem.Discount = discountPercentage;
                //newItem.IncludeNote = Note;
                //newItem.IncludeImage = Image;
                quoteItems.Add(newItem);
                rowCount++;
            }

            return quoteItems;
        }

        private void loadQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Int32 rowCount = OptionsDataGridView.Rows.GetRowCount(DataGridViewElementStates.Displayed);
            if (rowCount > 0)
            {
                if (MessageBox.Show("You already have items selected!\r\n\r\nDo you want to clear these items?", "Clear List", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    return;
                }

            }
            OptionsDataGridView.Rows.Clear();
            OptionsItemsPriceTextBox.Text = "";

            // Get Order Filename and Path
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\TecanFiles";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(Quote));
                System.IO.StreamReader file = new System.IO.StreamReader(openFileDialog1.FileName);
                Quote quote = new Quote();
                quote = (Quote)reader.Deserialize(file);

                // Seperate order name and date info, reformat date info
                string[] titleData = quote.QuoteTitle.Split('-');
                QuoteTitleTextBox.Text = titleData[0].Substring(0, titleData[0].Length);
                QuoteDateTimePicker.Value = quote.QuoteDate;

                String itemSAPID;
                String itemDescription;
                Decimal itemPrice;
                Decimal extendedPrice;
                Int32 itemQuantity;
                //Decimal itemDiscount;

                foreach (QuoteItems row in quote.Items)
                {
                    itemSAPID = row.SAPID;
                    itemDescription = row.Description;
                    itemPrice = row.Price;
                    itemQuantity = row.Quantity;
                    extendedPrice = itemPrice * itemQuantity;

                    OptionsDataGridView.Rows.Add(itemSAPID, itemDescription, itemPrice, itemQuantity, extendedPrice);
                }
                QuoteTabControl.SelectedTab = OptionTabPage;
                SumItems(OptionsDataGridView);

            }
        }

        private void myProfileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showUserProfileForm(false);
        }

        public void getUsersProfile()
        {
            String profileFile = @"c:\TecanFiles\" + "TecanConfig.cfg";
            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(Profile));
            System.IO.StreamReader file = new System.IO.StreamReader(profileFile);
            // Profile profile = new Profile();
            profile = (Profile)reader.Deserialize(file);
            file.Close();
            SalemansNameLabel.Text = "Welcome " + profile.Name;
            ProfileNameTextBox.Text = profile.Name;
            ProfileCompanyTextBox.Text = profile.Company;
            ProfilePhoneTextBox.Text = profile.Phone;
            ProfileEmailTextBox.Text = profile.Email;
            ProfileShippingAddressTextBox.Text = profile.ShippingAddress;
            ProfileCityTextBox.Text = profile.City;
            loadStateComboBox();
            ProfileStateComboBox.SelectedIndex = ProfileStateComboBox.FindStringExact(profile.State);
            ProfileZipcodeTextBox.Text = profile.Zipcode;
            ProfileTecanEmailTextBox.Text = profile.TecanEmail;

        }

        private void loadStateComboBox()
        {

            ProfileStateComboBox.Items.Add("AL");
            ProfileStateComboBox.Items.Add("AK");
            ProfileStateComboBox.Items.Add("AZ");
            ProfileStateComboBox.Items.Add("AR");
            ProfileStateComboBox.Items.Add("CA");
            ProfileStateComboBox.Items.Add("CO");
            ProfileStateComboBox.Items.Add("CT");
            ProfileStateComboBox.Items.Add("DE");
            ProfileStateComboBox.Items.Add("DC");
            ProfileStateComboBox.Items.Add("FL");
            ProfileStateComboBox.Items.Add("GA");
            ProfileStateComboBox.Items.Add("HI");
            ProfileStateComboBox.Items.Add("ID");
            ProfileStateComboBox.Items.Add("IL");
            ProfileStateComboBox.Items.Add("IN");
            ProfileStateComboBox.Items.Add("IA");
            ProfileStateComboBox.Items.Add("KS");
            ProfileStateComboBox.Items.Add("KY");
            ProfileStateComboBox.Items.Add("LA");
            ProfileStateComboBox.Items.Add("ME");
            ProfileStateComboBox.Items.Add("MD");
            ProfileStateComboBox.Items.Add("MA");
            ProfileStateComboBox.Items.Add("MI");
            ProfileStateComboBox.Items.Add("MN");
            ProfileStateComboBox.Items.Add("MS");
            ProfileStateComboBox.Items.Add("MO");
            ProfileStateComboBox.Items.Add("MT");
            ProfileStateComboBox.Items.Add("NE");
            ProfileStateComboBox.Items.Add("NV");
            ProfileStateComboBox.Items.Add("NH");
            ProfileStateComboBox.Items.Add("NJ");
            ProfileStateComboBox.Items.Add("NM");
            ProfileStateComboBox.Items.Add("NY");
            ProfileStateComboBox.Items.Add("NC");
            ProfileStateComboBox.Items.Add("ND");
            ProfileStateComboBox.Items.Add("OH");
            ProfileStateComboBox.Items.Add("OK");
            ProfileStateComboBox.Items.Add("OR");
            ProfileStateComboBox.Items.Add("PA");
            ProfileStateComboBox.Items.Add("RI");
            ProfileStateComboBox.Items.Add("SC");
            ProfileStateComboBox.Items.Add("SD");
            ProfileStateComboBox.Items.Add("TN");
            ProfileStateComboBox.Items.Add("TX");
            ProfileStateComboBox.Items.Add("UT");
            ProfileStateComboBox.Items.Add("VT");
            ProfileStateComboBox.Items.Add("VA");
            ProfileStateComboBox.Items.Add("WA");
            ProfileStateComboBox.Items.Add("WV");
            ProfileStateComboBox.Items.Add("WI");
            ProfileStateComboBox.Items.Add("WY");

        }

        private void saveUsersProfile()
        {
            // Save to Profile config file
            String profileFile = @"c:\TecanFiles\" + "TecanConfig.cfg";
            System.Xml.Serialization.XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(Profile));
            System.IO.StreamWriter file = new System.IO.StreamWriter(profileFile);
            writer.Serialize(file, profile);
            file.Close();
        }

        private void profileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showUserProfileForm(false);
        }

        //private void viewQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    string pdfTemplate = @"c:\temp\Tecan DiTi Systems Form - Ian2.pdf";
        //    this.Text += " - " + pdfTemplate;
        //    PdfReader pdfReader = new PdfReader(pdfTemplate);
        //    StringBuilder sb = new StringBuilder();
        //    foreach (DictionaryEntry de in pdfReader.AcroFields.Fields)
        //    {
        //        sb.Append(de.Key.ToString() + Environment.NewLine);
        //    }

        //    string fieldsFile = @"c:\temp\fields.txt";
        //    System.IO.File.WriteAllText(fieldsFile, sb.ToString());

        //    string newFile = @"c:\temp\output.pdf";
        //    PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
        //        newFile, FileMode.Create));
        //    AcroFields pdfFormFields = pdfStamper.AcroFields;

        //    // set form pdfFormFields
        //    pdfFormFields.SetField("Company", "The Company Name");
        //    pdfFormFields.SetField("Customer Name", "The Customer");
        //    pdfFormFields.SetField("Address", "123 Main Street");
        //    pdfFormFields.SetField("Phone", "919-555-1212");
        //    pdfFormFields.SetField("Email", "email@email.com");

        //    pdfFormFields.SetField("ReadersTotal", "100");
        //    pdfFormFields.SetField("WasherTotal", "200");
        //    pdfFormFields.SetField("QCKitTotal", "300");
        //    pdfFormFields.SetField("HPTotal", "400");
        //    pdfFormFields.SetField("TipsTotal", "1000");
        //    pdfFormFields.SetField("AppSuppTotal", "2000");
        //    pdfFormFields.SetField("ContractTotal", "3000");
        //    pdfFormFields.SetField("Price", "10000");
        //    pdfFormFields.SetField("Instrument Price", "10000");

        //    // flatten the form to remove editting options, set it to false
        //    // to leave the form open to subsequent manual edits
        //    pdfStamper.FormFlattening = false;
            
        //    // close the pdf
        //    pdfStamper.Close();
        //}

        private void clearQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OptionsDataGridView.Rows.Clear();
            QuoteTitleTextBox.Text = "";
            OptionsItemsPriceTextBox.Text = "";
        }

        private void updateDatabaseFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String distributionFolder = profile.DistributionFolder;
            Boolean fileFound;
            fileFound = copyDatabaseToWorkingFolder(distributionFolder);
            if (!fileFound)
            {
                MessageBox.Show("The Distribution Folder you selected in your profile does not contain the Parts List Database!\n\nPlease select a new folder");
                showUserProfileForm(true);
            }
            else
            {
                MessageBox.Show("New Parts List Database Loaded!");
                MainQuoteForm_Load(sender, e);
            }

        }

        private void updateActionStatus(string status)
        {
            if (this.IsHandleCreated)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    actionStatusLabel.Text = status;
                });
            }
            else
            {
                actionStatusLabel.Text = status;
            }
        }

        // This is the menu item for View Order (previously print order (quote))
        private void printQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            // Create the temp order textfile for printing
            string tempPath;
            tempPath = createTempFile(QuoteTitleTextBox.Text + ".txt", "txt");
            string fullTempPathName = tempPath + "\\" + QuoteTitleTextBox.Text + ".txt";

            //Open the file with the default associated application registered on the local machine
            Process.Start(fullTempPathName);

            //try
            //{
            //    streamToPrint = new StreamReader(fullTempPathName);
            //    try
            //    {
            //        printFont = new System.Drawing.Font("Arial", 10);
            //        PrintDocument pd = new PrintDocument();
            //        pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
            //        pd.Print();
            //    }
            //    finally
            //    {
            //        streamToPrint.Close();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            
        }

        // The PrintPage event is raised for each page to be printed.
        //private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        //{
        //    float linesPerPage = 0;
        //    float yPos = 0;
        //    int count = 0;
        //    float leftMargin = ev.MarginBounds.Left;
        //    float topMargin = ev.MarginBounds.Top;
        //    string line = null;

        //    // Calculate the number of lines per page.
        //    linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);

        //    // Print each line of the file.
        //    while (count < linesPerPage && ((line = streamToPrint.ReadLine()) != null))
        //    {
        //        yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));
        //        ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, yPos, new StringFormat());
        //        count++;
        //    }

        //    // If more lines exist, print another page.
        //    if (line != null)
        //        ev.HasMorePages = true;
        //    else
        //        ev.HasMorePages = false;
        //}

        private void sendQuoteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            // Create the temp order textfile for attachment (csv)
            string attachmentPath;
            attachmentPath = createTempFile(QuoteTitleTextBox.Text + ".csv", "csv");
            string fullAttachmentPathName = attachmentPath + "\\" + QuoteTitleTextBox.Text + ".csv";

            // Setup mail message
            MailAddress to = new MailAddress(ProfileTecanEmailTextBox.Text);
            MailAddress from = new MailAddress(ProfileEmailTextBox.Text);
            var mailMessage = new MailMessage(from, to);
            mailMessage.Subject = ProfileCompanyTextBox.Text + " Parts Order";
            mailMessage.Body = "Please process my attached order.";
            mailMessage.Attachments.Add(new Attachment(fullAttachmentPathName));

            var filename = attachmentPath + "\\mymessage.eml";

            //save the MailMessage to the filesystem
            mailMessage.Save(filename);

            //Open the file with the default associated application registered on the local machine
            Process.Start(filename);

        }

        private string createTempFile(String filespec, String format)
        {

            // Create the new file in temp directory
            String tempFilePath = @AppDomain.CurrentDomain.BaseDirectory.ToString() + "temp";
            System.IO.Directory.CreateDirectory(tempFilePath);

            // If temp directory current contains any files, delete them
            System.IO.DirectoryInfo tempFiles = new DirectoryInfo(tempFilePath);

            foreach (FileInfo file in tempFiles.GetFiles())
            {
                file.Delete();
            }

            String fullFilePathName = @tempFilePath + "\\" + filespec;

            string sData = "";
            sData += QuoteDateTimePicker.Text + "\n";
            sData += QuoteTitleTextBox.Text + "\n";
            sData += ProfileCompanyTextBox.Text + "\n";
            sData += ProfileNameTextBox.Text + "\n";
            sData += ProfileShippingAddressTextBox.Text + "\n";
            sData += ProfileCityTextBox.Text + " ";
            sData += ProfileStateComboBox.Text + " ";
            sData += ProfileZipcodeTextBox.Text + "\n";
            sData += ProfilePhoneTextBox.Text + "\n";
            sData += ProfileEmailTextBox.Text + "\n";
            sData += ProfileTecanEmailTextBox.Text + "\n\n";
            ArrayList quoteItems = new ArrayList();
            quoteItems = AddQuoteItems(OptionsDataGridView);
            string formatChar = "";
            if (format == "csv")
            {
                formatChar = ",";
            }
            else
            {
                formatChar = "\t";

            }

            foreach (QuoteItems row in quoteItems)
            {
                sData += row.SAPID + formatChar;

                if (format == "csv")
                {
                    sData += row.Description.Replace(",", "") + formatChar;
                }
                else
                {
                    if (row.Description.Length > 27)
                    {
                        sData += row.Description.Substring(0, 27) + formatChar;
                    }
                    else
                    {
                        sData += row.Description + formatChar;
                    }
                }
                sData += String.Format("{0:C}", row.Price) + formatChar;
                sData += row.Quantity + formatChar;
                sData += String.Format("{0:C}", row.Price * row.Quantity) + "\n";
            }
            if (format == "csv")
            {
                sData += "\n\n,,,," + OptionsItemsPriceTextBox.Text.Replace(",", "") + "\n";
            }
            else
            {
                sData += "\n\n" + OptionsItemsPriceTextBox.Text + "\n";
            }
            File.WriteAllText(fullFilePathName, sData);

            return tempFilePath;

        }
    }
}
