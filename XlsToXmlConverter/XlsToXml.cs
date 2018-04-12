using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XlsToXmlConverter.Classes;

namespace XlsToXmlConverter
{
    public partial class XlsToXml : Form
    {
        public XlsToXml()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (!openFileDialog.FileNames[0].Equals(""))
                {
                    ExcelHelper excelHelper = new ExcelHelper();
                    textBox1.Text = openFileDialog.FileNames[0];
                    dataGridView1.DataSource = excelHelper.ReadExcelFile(openFileDialog.FileNames[0]);
                }
            }
        }

        private void buttonSalvarXml_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelHelper excelHelper = new ExcelHelper();
                    string xml = excelHelper.GetXML(textBox1.Text);

                    using (StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName))
                    {
                        streamWriter.Write(xml);
                    }

                    MessageBox.Show("File saved successfully!");
                }
            }
        }
    }
}
