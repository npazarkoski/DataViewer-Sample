using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace SQL_Data_Viewer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox4.PasswordChar = '*';
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            hideUserAndPassword();
            listBox1.Visible = false;
            listBox1.Enabled = false;

            button2.Visible = false;
            button2.Enabled = false;

            dataGridView1.Visible = false;

            textBox5.Visible = false;

            disconnectBtn.Visible = false;
            disconnectBtn.Enabled = false;


            listBox2.Visible = false;
            listBox2.Enabled = false;

            label11.Visible = false;
            label12.Visible = false;

            hideFilters();
        }

        string cnString = "";
        SqlConnection sqlCon = new SqlConnection();

        DataTable columnDtbl = new DataTable();

        int columnType = -1;

        string filterStr = " 1=1 ";

        void hideUserAndPassword()
        {
            label3.Visible = false;
            textBox3.Visible = false;
            label4.Visible = false;
            textBox4.Visible = false;

            label3.Enabled = false;
            textBox3.Enabled = false;
            label4.Enabled = false;
            textBox4.Enabled = false;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                hideUserAndPassword();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                label3.Visible = true;
                textBox3.Visible = true;
                label4.Visible = true;
                textBox4.Visible = true;

                label3.Enabled = true;
                textBox3.Enabled = true;
                label4.Enabled = true;
                textBox4.Enabled = true;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox1.Text != "")
            {
                if (radioButton1.Checked)
                    cnString = @"Data Source=" + textBox1.Text + ";Initial Catalog=" + textBox2.Text + ";Integrated Security=True;";
                else
                    cnString = @"Data Source=" + textBox1.Text + ";Initial Catalog=" + textBox2.Text + ";User ID=" + textBox3.Text + ";Password=" + textBox4.Text + ";";

                sqlCon.ConnectionString = cnString;
                try
                {
                    if (sqlCon.State == System.Data.ConnectionState.Closed)
                    {
                        sqlCon.Open();
                        this.Size = new System.Drawing.Size(1700, 800);
                        onConnect();

                        SqlDataAdapter sqlda = new SqlDataAdapter("Select * from SYSOBJECTS WHERE  xtype = 'U'", sqlCon);
                        DataTable dtbl = new DataTable();
                        sqlda.Fill(dtbl);
                        listBox1.Items.Clear();
                        foreach (DataRow row in dtbl.Rows)
                        {
                            if (row[0].ToString() != "sysdiagrams")
                            {
                                listBox1.Items.Add(row[0].ToString());
                            }
                        }


                    }
                }
                catch {
                    MessageBox.Show("Connection could not be established with these input values !");
                }
            }
              
            else
            {
                if (textBox1.Text == "")
                    MessageBox.Show("Enter server name !");
                if (textBox2.Text == "")
                    MessageBox.Show("Enter database name !");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (sqlCon.State == System.Data.ConnectionState.Open)
                sqlCon.Close();
            cnString = "";


            this.Size = new System.Drawing.Size(757, 290);
            radioButton1.Enabled = true;
            radioButton2.Enabled = true;
            label1.Enabled = true;
            textBox1.Enabled = true;
            label2.Enabled = true;
            textBox2.Enabled = true;
            label3.Enabled = true;
            textBox3.Enabled = true;
            label4.Enabled = true;
            textBox4.Enabled = true;
            button1.Visible = true;
            button1.Enabled = true;

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";

            listBox1.Visible = false;
            listBox1.Enabled = false;

            disconnectBtn.Visible = false;
            disconnectBtn.Enabled = false;

            textBox5.Text = "";

            button2.Visible = false;
            button2.Enabled = false;

            dataGridView1.DataSource = null;
            dataGridView1.Visible = false;

            textBox5.Visible = false;

            listBox2.Visible = false;
            listBox2.Enabled = false;

            textBox9.Text = "";

            hideFilters();

            label11.Visible = false;
            label12.Visible = false;
        }


        void onConnect()
        {
            radioButton1.Enabled = false;
            radioButton2.Enabled = false;
            label1.Enabled = false;
            textBox1.Enabled = false;
            label2.Enabled = false;
            textBox2.Enabled = false;
            label3.Enabled = false;
            textBox3.Enabled = false;
            label4.Enabled = false;
            textBox4.Enabled = false;
            button1.Visible = false;
            button1.Enabled = false;

            listBox1.Visible = true;
            listBox1.Enabled = true;
            listBox1.Items.Clear();

            button2.Visible = true;
            button2.Enabled = true;

            dataGridView1.Visible = true;

            textBox5.Visible = true;

            disconnectBtn.Visible = true;
            disconnectBtn.Enabled = true;

            listBox2.Visible = true;
            listBox2.Enabled = true;
            listBox2.Items.Clear();

            label11.Visible = true;
            label12.Visible = true;
        }

        private String getPrimaryKey(string table)
        {
            string primary = "";
            SqlDataAdapter sqlda = new SqlDataAdapter("EXEC sp_pkeys " + table, sqlCon);
            DataTable dtbl = new DataTable();
            sqlda.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                if (dtbl.Rows.IndexOf(row) != 0)
                    primary += ", ";
                primary += row[3].ToString();
            }
            return primary;
        }
        private String getForeignKey(string table)
        {
            string foreign = "";
            SqlDataAdapter sqlda = new SqlDataAdapter("EXEC sp_fkeys " + table, sqlCon);
            DataTable dtbl = new DataTable();
            sqlda.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                foreign += dtbl.Rows.IndexOf(row) + 1 + ". " + "primary key column:" + row[3].ToString() + " , foreign key table: " + row[6].ToString() + " , foreign key column:" + row[7].ToString() + Environment.NewLine;
            }
            return foreign;
        }

        private string getNumOfRows()
        {
            SqlCommand command = new SqlCommand("SELECT @@ROWCOUNT FROM " + listBox1.GetItemText(listBox1.SelectedItem), sqlCon);

            int i = 0;

            SqlDataReader reader = command.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    i += 1;
                }
            }

            reader.Close();
            return Convert.ToString(i);
        }

        private void populateDataGrid()
        {
            SqlDataAdapter sqlda = new SqlDataAdapter("Select * from " + listBox1.GetItemText(listBox1.SelectedItem), sqlCon);
            DataTable dtbl = new DataTable();
            sqlda.Fill(dtbl);
            dataGridView1.DataSource = dtbl;

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox5.Text = "";
            textBox5.Text += "Primary key columns: " + getPrimaryKey(listBox1.GetItemText(listBox1.SelectedItem)) + Environment.NewLine + Environment.NewLine;
            textBox5.Text += "Foreign key relationships: \n" + getForeignKey(listBox1.GetItemText(listBox1.SelectedItem)) + Environment.NewLine;
            textBox5.Text += "Number of rows: " + getNumOfRows();

            populateDataGrid();

            getColumnNames();

            textBox9.Text = "";

            filterStr = " 1=1 ";

            hideFilters();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count > 0)
            {
                try
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                    Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                    app.Visible = true;

                    worksheet = workbook.Sheets[1];

                    worksheet = workbook.ActiveSheet;

                    worksheet.Name = "Exported from gridview";

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {

                        worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;

                    }

                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {

                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {

                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                        }

                    }
                    app.Columns.AutoFit();
                    app.Visible = true;

                    MessageBox.Show("Exported");
                }
                catch { }
            }

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            hideFilters();
            button3.Visible = true;
            button3.Enabled = true;

            button4.Visible = true;
            button4.Enabled = true;

            button5.Visible = true;
            button5.Enabled = true;

            textBox9.Visible = true;

            columnType = checkColumnType();
            if (columnType == 1 ^ columnType == 3)
            {
                label5.Visible = true;
                label6.Visible = true;


                textBox6.Visible = true;
                textBox7.Visible = true;
                textBox6.Enabled = true;
                textBox7.Enabled = true;

                if (columnType == 1)
                {
                    label5.Text = "From";
                    label6.Text = "To";

                }
                else
                {
                    label5.Text = "Starts with";
                    label6.Text = "Ends with";

                    label8.Visible = true;
                    textBox8.Visible = true;
                    textBox8.Enabled = true;
                }
            }
            else if (columnType == 2)
            {
                label9.Visible = true;
                label10.Visible = true;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
            }

        }

        private void getColumnNames()
        {
            columnDtbl.Clear();
            SqlDataAdapter sqlda = new SqlDataAdapter("SELECT COLUMN_NAME , NUMERIC_PRECISION, DATETIME_PRECISION,CHARACTER_SET_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '"
            + listBox1.GetItemText(listBox1.SelectedItem) + "'", sqlCon);
            sqlda.Fill(columnDtbl);
            listBox2.Items.Clear();
            foreach (DataRow row in columnDtbl.Rows)
            {
                listBox2.Items.Add(row[0].ToString());
            }

        }

        private int checkColumnType()
        {
            foreach (DataRow row in columnDtbl.Rows)
            {

                if (listBox2.GetItemText(listBox2.SelectedItem) == row[0].ToString())
                {
                    if (row[1].ToString() != "")
                        return 1;
                    else if (row[2].ToString() != "")
                        return 2;
                    else if (row[3].ToString() != "")
                        return 3;
                }
            }
            return -1;
        }

        private void hideFilters()
        {
            label5.Visible = false;
            label6.Visible = false;
            label8.Visible = false;
            label9.Visible = false;
            label10.Visible = false;

            textBox6.Visible = false;
            textBox7.Visible = false;
            textBox8.Visible = false;

            textBox6.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;

            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;

            button3.Visible = false;
            button3.Enabled = false;

            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Visible = false;

            button4.Visible = false;
            button4.Enabled = false;

            button5.Visible = false;
            button5.Enabled = false;

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (columnType == 2)
            {    
                if (dateTimePicker1.Value.Date>dateTimePicker2.Value.Date){
                    MessageBox.Show("Wrong dates !"); }
                else {
                filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + ">= '" + dateTimePicker1.Value.Date + "'";
                filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + "<= '" + dateTimePicker2.Value.Date + "'";
            }
            }
            else
            {
                if (textBox6.Text != "")
                {
                    if (columnType == 1)
                    {
                        filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + ">=" + textBox6.Text;
                        textBox6.Text = "";
                    }
                    else if (columnType == 3)
                    {
                        filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + " LIKE '" + textBox6.Text + "%'";
                        textBox6.Text = "";
                    }
                }
                if (textBox7.Text != "")
                {

                    if (columnType == 1)
                    {
                        filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + "<=" + textBox7.Text;
                        textBox7.Text = "";
                    }
                    else if (columnType == 3)
                    {
                        filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + " LIKE '%" + textBox7.Text + "'";
                        textBox7.Text = "";
                    }


                }
                if (textBox8.Text != "" && columnType == 3)
                {
                    filterStr += " AND " + listBox2.GetItemText(listBox2.SelectedItem) + " LIKE '%" + textBox8.Text + "%'";
                    textBox8.Text = "";
                }
            }

            textBox9.Text = filterStr;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            filterStr = " 1=1 ";
            textBox9.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter sqlda = new SqlDataAdapter("Select * from " + listBox1.GetItemText(listBox1.SelectedItem) + " Where " + filterStr, sqlCon);
                DataTable dtbl = new DataTable();
                sqlda.Fill(dtbl);
                dataGridView1.DataSource = dtbl;
            }
            catch
            {
                MessageBox.Show("Wrong input parameters ! ");
            }

        }


    }
}
