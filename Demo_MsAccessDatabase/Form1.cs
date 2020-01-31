using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace Demo_MsAccessDatabase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ReadData();
            dataGridView1.DataSource = GetData(
                                    "|DataDirectory|\\DatabaseFiles\\demo.accdb",
                                    "SELECT * FROM Students;");
        }
        private void ReadData()
        {
            //prepare
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString =
                "Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=|DataDirectory|\\DatabaseFiles\\demo.accdb";
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            //SELECT * FROM table_name;
            command.CommandText = "SELECT * FROM Students;";//Sql
            DataTable table_student = new DataTable("Students");
            //connect
            connection.Open();
            //execute SQL
            OleDbDataReader data_reader = command.ExecuteReader();
            table_student.Load(data_reader);//while loop
            //close connection
            connection.Close();
            //connect table_student to DataGridView
            dataGridView1.DataSource = table_student;//generate column
                                                     //load data
        }
        private DataTable GetData(string ms_access_file, string sql_select)
        {
            //prepare
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString =
                "Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=" + ms_access_file;
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            //SELECT * FROM table_name;
            command.CommandText = sql_select;//Sql
            DataTable table_student = new DataTable("Students");
            //connect
            connection.Open();
            //execute SQL
            OleDbDataReader data_reader = command.ExecuteReader();
            table_student.Load(data_reader);//while loop
            //close connection
            connection.Close();
            //connect table_student to DataGridView
            return table_student;//generate column
                                 //load data
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetData(
                                    "|DataDirectory|\\DatabaseFiles\\demo.accdb",
                                    "SELECT * FROM Students;");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string ms_access_file = "|DataDirectory|\\DatabaseFiles\\demo.accdb";
            if (editing_row == null)
            {
                string sql_insert = "INSERT INTO " +
                    " Students(Id, StuddentName, Gender, Phone, Email) " +
                    " VALUES (" +
                    txtId.Text + "," +
                    "'" + txtName.Text + "'," +
                    "'" + cboGender.Text + "'," +
                    "'" + txtPhone.Text + "'," +
                    "'" + txtEmail.Text + "'" +
                    ")";
                int result_count = RunSqlCommand(ms_access_file, sql_insert);
                //ClearForm();
            }
            else
            {
                //update database
                string sql_update = "UPDATE Students SET " +
                    "StuddentName = '" + txtName.Text.Trim() + "', " +
                    "Gender = '" + cboGender.Text.Trim() + "', " +
                    "Phone = '" + txtPhone.Text.Trim() + "', " +
                    "Email = '" + txtEmail.Text.Trim() + "'" +
                    " Where Id = " + txtId.Text;
                RunSqlCommand(ms_access_file, sql_update);
                //update datagridview
                editing_row.Cells[1].Value = txtName.Text.Trim();
                editing_row.Cells[2].Value = cboGender.Text.Trim();
                editing_row.Cells[3].Value = txtPhone.Text.Trim();
                editing_row.Cells[4].Value = txtEmail.Text.Trim();
                editing_row = null;
                //ClearForm
            }
        }
        private int RunSqlCommand(string ms_access_file, string sql)
        {
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString =
                "Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=" + ms_access_file;//Error1
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            command.CommandText = sql;//Error2
            //connect
            connection.Open();//Error1
            //run SQL statement
            int result_count = command.ExecuteNonQuery();//run//Error2
            connection.Close();
            return result_count;
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            MessageBox.Show(e.Row.Index.ToString());
        }

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {

        }

        private void dataGridView1_UserDeletingRow_1(object sender, DataGridViewRowCancelEventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you to delete?",
                "Confirms",
                MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                //delete from database
                string ms_access_file = "|DataDirectory|\\DatabaseFiles\\demo.accdb";
                string id = e.Row.Cells[0].Value.ToString();//
                string sql_delete = "DELETE FROM Students WHERE Id = " + id;
                int result_count = RunSqlCommand(ms_access_file, sql_delete);
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string ms_access_file = "|DataDirectory|\\DatabaseFiles\\demo.accdb";
                string id = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();//
                string sql_delete = "DELETE FROM Students WHERE Id = " + id;
                int result_count = RunSqlCommand(ms_access_file, sql_delete);
                if (result_count > 0)
                {
                    dataGridView1.Rows.Remove(dataGridView1.SelectedRows[0]);
                }
            }
        }
        DataGridViewRow editing_row;
        private void Edit_Click(object sender, EventArgs e)
        {
            editing_row = dataGridView1.SelectedRows[0];
            txtId.Text = editing_row.Cells[0].Value.ToString();
            txtName.Text = editing_row.Cells[1].Value.ToString();
            cboGender.Text = editing_row.Cells[2].Value.ToString();
            txtPhone.Text = editing_row.Cells[3].Value.ToString();
            txtEmail.Text = editing_row.Cells[4].Value.ToString();
            //
            txtId.ReadOnly = true;
            txtName.SelectAll();
            txtName.Focus();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            bool conversion_result = int.TryParse(txtSearch.Text.Trim(), out int id);
            string sql_select = string.Empty;
            if (conversion_result == true)
            {
                sql_select = "SELECT * FROM Students WHERE Id = " + txtSearch.Text.Trim() + ";";
            }
            else
            {
                sql_select = "SELECT * FROM Students WHERE StuddentName LIKE '%" + txtSearch.Text.Trim() + "%' OR Phone = '" + txtSearch.Text.Trim() + "';";
            }
            string ms_access_file = "|DataDirectory|\\DatabaseFiles\\demo.accdb";
            dataGridView1.DataSource = GetData(ms_access_file, sql_select);
        }
    }
}
