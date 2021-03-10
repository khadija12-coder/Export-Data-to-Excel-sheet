using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public SqlConnection BD = new SqlConnection();
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
           

        }

        private void Form1_Load(object sender, EventArgs e)
        {

           
            
            
    
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (BD.State == ConnectionState.Closed)
            {
                BD.ConnectionString = (@"Data Source=DESKTOP-TTQNQUI\SQLEXPRESSK;Initial Catalog=emsi;Integrated Security=True");
                BD.Open();
                MessageBox.Show("Ouverte");
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            if (BD.State == ConnectionState.Open)
            {

                BD.Close();
                MessageBox.Show("Fermee");
            }
        }
        // mode connecte
        public SqlCommand SelectCommande = new SqlCommand();
        public SqlDataReader reader;
        
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            SelectCommande.Connection = BD;
            SelectCommande.CommandText = "Select * From Produit";
            reader = SelectCommande.ExecuteReader();
            while (reader.Read())
            {
                dataGridView1.Rows.Add(reader[0], reader[1], reader[2], reader[3]);
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            // Create an instance of the class that exports Excel files
            Excel.Application excel = new Excel.Application();
           // Makes Excel visible
            excel.Visible = true;
            object Missing = Type.Missing;
            // create a new Excel Workbook
            Workbook workbook = excel.Workbooks.Add(Missing);
           // After creating the new Workbook, next step is to create a new worksheet
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            
            int StartCol = 1;
            int StartRow = 1;

            //Adding DataGridView header
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                            Range myRange = (Range)sheet1.Cells[StartRow,StartCol+j] ;
                             myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                             // Change the header color 
                             myRange.Font.Color = Color.Red;
            
            }
            StartRow++;

            //Adding DataGridView in Cells
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 =  dataGridView1[j, i].Value;
                    
                }
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
