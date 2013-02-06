using BitCollectors.UIAutomationLib.Entities;
using BitCollectors.UIAutomationLib.Helpers;
using System;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace BitCollectors.UIAutomationLib.TestHarness
{
    public partial class lblAutomationFile : Form
    {
        private string automationFile = null;

        public lblAutomationFile()
        {
            InitializeComponent();
        }

        private void btnRunTest_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(automationFile))
            {
                MessageBox.Show("Please choose an automation XML file", "No automation file specified", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            StringDictionary macros = new StringDictionary();
            macros.Add("{USERTEXT}", textBox1.Text);

            UIAutomation automationConfig = XmlHelper.ProcessXmlFile(automationFile, macros);

            if (automationConfig != null)
            {
                Win32Helper win32 = new Win32Helper();
                win32.ExecuteAutomation(automationConfig);
            }
        }

        private void btnOpenAutomationFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = false;
            openFileDialog1.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                automationFile = openFileDialog1.FileName;
            }
        }
    }
}
