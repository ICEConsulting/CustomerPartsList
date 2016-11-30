using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Tecan_Parts
{
    public partial class ProfileForm : Form
    {

        MainQuoteForm mainForm;
        Boolean doInitialization;
        public void SetForm1Instance(MainQuoteForm inst)
        {
            mainForm = inst;
        }

        public ProfileForm(Boolean NeedsBD)
        {
            InitializeComponent();
            doInitialization = NeedsBD;
            loadStateComboBox();
            String profileFile = @"c:\TecanFiles\" + "TecanConfig.cfg";
            if (File.Exists(profileFile))
                {
                System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(Profile));
                System.IO.StreamReader file = new System.IO.StreamReader(profileFile);
                Profile profile = new Profile();
                profile = (Profile)reader.Deserialize(file);
                file.Close();
                ProfileNameTextBox.Text = profile.Name;
                ProfileCompanyTextBox.Text = profile.Company;
                ProfilePhoneTextBox.Text = profile.Phone;
                ProfileEmailTextBox.Text = profile.Email;
                ProfileShippingAddressTextBox.Text = profile.ShippingAddress;
                ProfileCityTextBox.Text = profile.City;
                ProfileStateComboBox.SelectedIndex = ProfileStateComboBox.FindStringExact(profile.State);
                ProfileZipcodeTextBox.Text = profile.Zipcode;
                ProfileTecanEmailTextBox.Text = profile.TecanEmail;
                ProfileDistributionFolderTextBox.Text = profile.DistributionFolder;
            }
        }

        private void loadStateComboBox()
        {
            ProfileStateComboBox.Items.Add("AK");
            ProfileStateComboBox.Items.Add("AL");
            ProfileStateComboBox.Items.Add("AR");
            ProfileStateComboBox.Items.Add("AZ");
            ProfileStateComboBox.Items.Add("CA");
            ProfileStateComboBox.Items.Add("CO");
            ProfileStateComboBox.Items.Add("CT");
            ProfileStateComboBox.Items.Add("DC");
            ProfileStateComboBox.Items.Add("DE");
            ProfileStateComboBox.Items.Add("FL");
            ProfileStateComboBox.Items.Add("GA");
            ProfileStateComboBox.Items.Add("HI");
            ProfileStateComboBox.Items.Add("IA");
            ProfileStateComboBox.Items.Add("ID");
            ProfileStateComboBox.Items.Add("IL");
            ProfileStateComboBox.Items.Add("IN");
            ProfileStateComboBox.Items.Add("KS");
            ProfileStateComboBox.Items.Add("KY");
            ProfileStateComboBox.Items.Add("LA");
            ProfileStateComboBox.Items.Add("MA");
            ProfileStateComboBox.Items.Add("MD");
            ProfileStateComboBox.Items.Add("ME");
            ProfileStateComboBox.Items.Add("MI");
            ProfileStateComboBox.Items.Add("MN");
            ProfileStateComboBox.Items.Add("MO");
            ProfileStateComboBox.Items.Add("MS");
            ProfileStateComboBox.Items.Add("MT");
            ProfileStateComboBox.Items.Add("NC");
            ProfileStateComboBox.Items.Add("ND");
            ProfileStateComboBox.Items.Add("NE");
            ProfileStateComboBox.Items.Add("NH");
            ProfileStateComboBox.Items.Add("NJ");
            ProfileStateComboBox.Items.Add("NM");
            ProfileStateComboBox.Items.Add("NV");
            ProfileStateComboBox.Items.Add("NY");
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
            ProfileStateComboBox.Items.Add("VA");
            ProfileStateComboBox.Items.Add("VT");
            ProfileStateComboBox.Items.Add("WA");
            ProfileStateComboBox.Items.Add("WI");
            ProfileStateComboBox.Items.Add("WV");
            ProfileStateComboBox.Items.Add("WY");
        }

        private void BrowseDistributionFolderButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.Description = "Please select your Quote Database Distribution Folder";
            folderBrowserDialog1.ShowNewFolderButton = false;

            String sourcePath = "";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                sourcePath = folderBrowserDialog1.SelectedPath;
            }
            ProfileDistributionFolderTextBox.Text = sourcePath;
        }

        private void ProfileSaveButton_Click(object sender, EventArgs e)
        {
            String errorMessage = "All fields must be entered!\n\nPlease enter \\ correct the following information.\n\n";
            Boolean profileError = false;
            RegexUtilities util;
            if (ProfileNameTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Your Full Name.\n\n";
            }
            if (ProfileCompanyTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Your Company's Name.\n\n";
            }
            if (ProfilePhoneTextBox.Text == "" || ProfilePhoneTextBox.Text.Length < 12)
            {
                profileError = true;
                errorMessage = errorMessage + "Your Phone Number.\n\n";
            }
            util = new RegexUtilities();
            if (ProfileEmailTextBox.Text == "" || !util.IsValidEmail(ProfileEmailTextBox.Text))
            {
                profileError = true;
                errorMessage = errorMessage + "Your Email Address.\n\n";
            }
            if (ProfileShippingAddressTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Your Shipping.\n\n";
            }
            if (ProfileCityTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Your City.\n\n";
            }
            if (ProfileZipcodeTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Your Zipcode.\n\n";
            }
            util = new RegexUtilities();
            if (ProfileTecanEmailTextBox.Text == "" || !util.IsValidEmail(ProfileTecanEmailTextBox.Text))
            {
                profileError = true;
                errorMessage = errorMessage + "Your Tecan Email Address.\n\n";
            }
            if (ProfileDistributionFolderTextBox.Text == "")
            {
                profileError = true;
                errorMessage = errorMessage + "Tecan Database Distribution Folder Location.\n\n";
            }

            if (profileError)
            {
                MessageBox.Show(errorMessage);
                // return;
            }
            else
            {
                Profile profile = new Profile();

                profile.Name = ProfileNameTextBox.Text;
                profile.Company = ProfileCompanyTextBox.Text;
                profile.Phone = ProfilePhoneTextBox.Text;
                profile.Email = ProfileEmailTextBox.Text;
                profile.ShippingAddress = ProfileShippingAddressTextBox.Text;
                profile.City = ProfileCityTextBox.Text;
                profile.State = ProfileStateComboBox.Text;
                profile.Zipcode = ProfileZipcodeTextBox.Text;
                profile.TecanEmail = ProfileTecanEmailTextBox.Text;
                profile.DistributionFolder = ProfileDistributionFolderTextBox.Text;

                // Save to Profile config file
                String tecanFilesFilePath = @"c:\TecanFiles";
                System.IO.Directory.CreateDirectory(tecanFilesFilePath);

                String profileFile = @"c:\TecanFiles\" + "TecanConfig.cfg";
                System.Xml.Serialization.XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(Profile));
                System.IO.StreamWriter file = new System.IO.StreamWriter(profileFile);
                writer.Serialize(file, profile);
                file.Close();
                this.Close();
                
                if (doInitialization)
                {
                    Boolean fileFound;
                    fileFound = mainForm.copyDatabaseToWorkingFolder(profile.DistributionFolder);
                    if(!fileFound)
                    {
                        MessageBox.Show("The Distribution Folder you selected does not contain the Parts List Database!\n\nPlease select a new folder");
                        mainForm.showUserProfileForm(true);
                        //ProfileForm profileForm = new ProfileForm(true);
                        //profileForm.Show();
                        //Application.OpenForms["ProfileForm"].BringToFront();
                    }
                    else
                    {
                        mainForm.getUsersProfile();
                        mainForm.MainQuoteForm_Load(sender, e);
                    }
                }
                else
                {
                    mainForm.getUsersProfile();
                }
            }

        }

        private void ProfilePhoneTextBox_TextChanged(object sender, EventArgs e)
        {
            if (ProfilePhoneTextBox.Text.Length == 3)
                ProfilePhoneTextBox.Text = ProfilePhoneTextBox.Text + "-";
            if (ProfilePhoneTextBox.Text.Length == 7)
                ProfilePhoneTextBox.Text = ProfilePhoneTextBox.Text + "-";
                ProfilePhoneTextBox.SelectionStart = ProfilePhoneTextBox.Text.Length;

        }

    }
}
