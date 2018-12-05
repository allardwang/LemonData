using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace EasyXML
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            bool isForbid, isAutoCheck, isAutoClose;
            ThisAddIn.RedConfig(out isForbid, out isAutoCheck, out isAutoClose);
            checkBox1.Checked = isForbid;
            checkBox2.Checked = isAutoCheck;
            checkBox3.Checked = isAutoClose;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.OpenFile();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.SaveFile();
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            SetConfig();
        }

        private void checkBox2_Click(object sender, RibbonControlEventArgs e)
        {
            SetConfig();
        }

        private void checkBox3_Click(object sender, RibbonControlEventArgs e)
        {
            SetConfig();
        }

        private void SetConfig()
        {
            ThisAddIn.SetConfig(checkBox1.Checked, checkBox2.Checked, checkBox3.Checked);
        }

    }
}
