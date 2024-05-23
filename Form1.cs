using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        DataGridViewRow selectedRow;

        public Form1()
        {
            InitializeComponent();
        }



        private void addIameColumn()
        {
            
            DataGridViewImageColumn imgCol = new DataGridViewImageColumn();
            imgCol.Name = "Image";
            dataGridView1.Columns.Add(imgCol);
            
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewImageCell imageCell = row.Cells["Image"] as DataGridViewImageCell;

                if (row.Cells["ImagePath"].Value  ==null|| row.Cells["ImagePath"].Value.Equals(""))
                {
                    continue;
                }
                
                imageCell.ImageLayout = DataGridViewImageCellLayout.Stretch;
                imageCell.Value = Image.FromFile(row.Cells[7].Value.ToString());
                
            }
        }


        public OleDbConnection GetConnection()
        {
            String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db/STUDENTS1.mdb";
            OleDbConnection conn = new OleDbConnection(connectionString);  // אוביקט ה"חיבור"
            conn.Open();    // DataBaseפתיחת החיבור ל-
            return conn;
        }


        private void Insret_Click(object sender, EventArgs e)
        {
            // Validate input data
            if (!int.TryParse(textBox1.Text, out int sid))
            {
                MessageBox.Show("Invalid Student ID. Please enter a valid integer.");
                return;
            }

            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Student Name cannot be empty. Please enter a name.");
                return;
            }

            if (string.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("Student Email cannot be empty. Please enter an email address.");
                return;
            }

            if (!int.TryParse(textBox4.Text, out int snumber))
            {
                MessageBox.Show("Invalid Phone Number. Please enter a valid integer.");
                return;
            }

            if (string.IsNullOrEmpty(textBox5.Text))
            {
                MessageBox.Show("Student Language cannot be empty. Please enter a language.");
                return;
            }

            if (string.IsNullOrEmpty(textBox6.Text))
            {
                MessageBox.Show("Student Country cannot be empty. Please enter a country.");
                return;
            }

            if (string.IsNullOrEmpty(textBox7.Text))
            {
                MessageBox.Show("Student Gender cannot be empty. Please enter a gender.");
                return;
            }

            if (string.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show("Student Image Path cannot be empty. Please enter an image path.");
                return;
            }
            OleDbConnection conn = GetConnection();  // אוביקט ה"חיבור דרך הפונקציה"

            sid = int.Parse(textBox1.Text);
            string sname = textBox2.Text;
            string semail = textBox3.Text;
            snumber = int.Parse(textBox4.Text);
            string slanguage = textBox5.Text;
            string scountry = textBox6.Text;
            string sgender = textBox7.Text;
            string sImagePath = textBox8.Text;

            string query =
                 "INSERT INTO students (ID,name,Email,PhoneNumber,Languag,Country,Gender,ImagePath) " +
                 "VALUES(\'" + sid + "\', \'" + sname + "\', \'" + semail + "\', \'" + snumber + "\', \'" + slanguage + "\', \'" + scountry + "\', \'" + sgender + "\',\'" + sImagePath + "');";
            MessageBox.Show(query);
            OleDbCommand cmd = new OleDbCommand(query, conn);
            int rows = cmd.ExecuteNonQuery();
            MessageBox.Show(rows + "added");

            Clear_dgv_Click();
            Refresh_dgv_Click();
            
            conn.Close();
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = GetConnection();

            int id = int.Parse(textBox1.Text);
            string query = "Delete from  students where id=" + id;

            OleDbCommand cmd = new OleDbCommand(query, conn);       // יצירת אובייקט כדי להריץ את השאילתה
            DialogResult dialogResult = MessageBox.Show(" Are you Sure?", "Conrifm", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                int rows = cmd.ExecuteNonQuery();     // הרצת שאילתה (שינויים)
                MessageBox.Show(rows + "added");
            }
            else if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("Canceled");
            }

            conn.Close();
            Clear_dgv_Click();
            Refresh_dgv_Click();
        }

        private void Update_Click(object sender, EventArgs e)
        {
            

            OleDbConnection conn = GetConnection();

            int sid = int.Parse(textBox1.Text);
            string sname = textBox2.Text;
            string semail = textBox3.Text;
            int snumber = int.Parse(textBox4.Text);
            string slanguage = textBox5.Text;
            string scountry = textBox6.Text;
            string sgender = textBox7.Text;
            string sImagePath = textBox8.Text;

            string query =
                "UPDATE students " +
                "SET name = '" + sname + "',Email= '" + semail + "',PhoneNumber= '" + snumber + "',Languag= '" + slanguage + "',Country= '" + scountry + "',Gender= '" + sgender + "',ImagePath='" + sImagePath + "'" +
                "WHERE ID = " + sid + ";";

            MessageBox.Show(query);

            OleDbCommand cmd = new OleDbCommand(query, conn);    // יצירת אובייקט כדי להריץ את השאילתה
            int rows = cmd.ExecuteNonQuery();
            MessageBox.Show(rows + "added");

            conn.Close();
            Clear_dgv_Click();
            Refresh_dgv_Click();

        }

        private void Clear_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = GetConnection();

            string query = "Delete from  students ";

            OleDbCommand cmd = new OleDbCommand(query, conn);       // יצירת אובייקט כדי להריץ את השאילתה
            DialogResult dialogResult = MessageBox.Show(" Are you Sure?", "Conrifm", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                int rows = cmd.ExecuteNonQuery();     // הרצת שאילתה (שינויים)
                MessageBox.Show(rows + "added");
            }
            else if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("Canceled");
            }

            conn.Close();
            Clear_dgv_Click();
            Refresh_dgv_Click();
        }

        private void Clear_dgv_Click(object sender=null, EventArgs e=null)
        {
            dataGridView1.Columns.Clear();
           
        }

        private void Refresh_dgv_Click(object sender=null, EventArgs e=null)
        {
            dataGridView1.Columns.Clear();
            string sql = "SELECT * FROM students";
            OleDbConnection connection = GetConnection();
            OleDbDataAdapter dataadapter = new OleDbDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            dataadapter.Fill(ds, "STUDENTS_table");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "STUDENTS_table";
            connection.Close();
            addIameColumn();
        }

        private void ViewSelectedRowsonNewForm_Click(object sender, EventArgs e)
        {
            Form2 fr2 = new Form2();
            if (selectedRow != null)
            {
                fr2.textBox1.Text = selectedRow.Cells[0].Value.ToString();
                fr2.textBox2.Text = selectedRow.Cells[1].Value.ToString();
                fr2.textBox3.Text = selectedRow.Cells[2].Value.ToString();
                fr2.textBox4.Text = selectedRow.Cells[3].Value.ToString();
                fr2.textBox5.Text = selectedRow.Cells[4].Value.ToString();
                fr2.textBox6.Text = selectedRow.Cells[5].Value.ToString();
                fr2.textBox7.Text = selectedRow.Cells[6].Value.ToString();
                fr2.textBox8.Text = selectedRow.Cells[7].Value.ToString();
            }

            fr2.Show();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv != null && dgv.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dgv.SelectedRows[0];
                if (row != null)
                {
                    selectedRow = row;
                    textBox1.Text = row.Cells[0].Value.ToString();
                    textBox2.Text = row.Cells[1].Value.ToString();
                    textBox3.Text = row.Cells[2].Value.ToString();
                    textBox4.Text = row.Cells[3].Value.ToString();
                    textBox5.Text = row.Cells[4].Value.ToString();
                    textBox6.Text = row.Cells[5].Value.ToString();
                    textBox7.Text = row.Cells[6].Value.ToString();
                    textBox8.Text = row.Cells[7].Value.ToString();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // הצגת הטבלה 
            string sql = "SELECT * FROM students";
            OleDbConnection connection = GetConnection();
            OleDbDataAdapter dataadapter = new OleDbDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            dataadapter.Fill(ds, "STUDENTS_table");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "STUDENTS_table";
            connection.Close();
            addIameColumn();
        }
    }
}
