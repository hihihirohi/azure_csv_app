using System;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using CsvHelper;
using CsvHelper.Configuration;


namespace AzureCsvApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            InportData.Inport();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            SumPrice.Output();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            LeadPrice.Output();
        }
    }
}