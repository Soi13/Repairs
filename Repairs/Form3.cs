using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;

namespace Repairs
{
    public partial class Form3 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=REPAIRS;Data Source=T1212-W00079\MSSQLSERVER2012");
            
        SqlDataAdapter da;
        Form2 form2 = new Form2();

        public Form3()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        private void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["STATION_NAME"].HeaderText = "Наименование станции";
            dataGridView1.Columns["STATION_NAME"].Width = 100;
            dataGridView1.Columns["EQUIPMENT_NAME"].HeaderText = "Наименование оборудования";
            dataGridView1.Columns["EQUIPMENT_NAME"].Width = 100;
            dataGridView1.Columns["TYPE_WORK"].HeaderText = "Вид работ";
            dataGridView1.Columns["TYPE_WORK"].Width = 70;
            dataGridView1.Columns["TYPE_REPAIR"].HeaderText = "Вид ремонта";
            dataGridView1.Columns["TYPE_REPAIR"].Width = 70;
            dataGridView1.Columns["PERIOD_REPAIR_BEGIN"].HeaderText = "Начало ремонта";
            dataGridView1.Columns["PERIOD_REPAIR_BEGIN"].Width = 100;
            dataGridView1.Columns["PERIOD_REPAIR_END"].HeaderText = "Окончание ремонта";
            dataGridView1.Columns["PERIOD_REPAIR_END"].Width = 100;
            dataGridView1.Columns["MOVE_PERIOD_REPAIR"].HeaderText = "Перенос срока окончания";
            dataGridView1.Columns["MOVE_PERIOD_REPAIR"].Width = 100;
            dataGridView1.Columns["EXPLAIN_MOVE_REPAIR"].HeaderText = "Обоснование переноса срока";
            dataGridView1.Columns["EXPLAIN_MOVE_REPAIR"].Width = 300;
            dataGridView1.Columns["VOLUME"].HeaderText = "Плановый объем";
            dataGridView1.Columns["VOLUME"].Width = 200;
            dataGridView1.Columns["VIPOLNENO_TEK_DATA"].HeaderText = "Выполнено на тек. дату";
            dataGridView1.Columns["VIPOLNENO_TEK_DATA"].Width = 200;
            dataGridView1.Columns["PROCENT_VIPOLNEN"].HeaderText = "Процент выполнения";
            dataGridView1.Columns["PROCENT_VIPOLNEN"].Width = 100;
            dataGridView1.Columns["ID_USER"].Visible = false;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["PERIOD"].Visible = false;
            dataGridView1.Columns["STATUS_UTVERJDEN"].Visible = false;
        }
        //////////////////////////////////////////////////////
     

        private void button1_Click(object sender, EventArgs e)
        {


            if ((checkBox1.Checked == false) && (checkBox2.Checked == false) && (checkBox3.Checked == false) && (checkBox4.Checked == false) && (checkBox5.Checked == false))
            {
                MessageBox.Show("Не выбрано ни одного параметра!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SystemSounds.Beep.Play();
                return;
            }
            
            ////////////////////////////
            if (checkBox1.Checked==true)
            {
                if (comboBox1.Text.Length == 0)
                   {
                       MessageBox.Show("Не заполнено поле \"Наименование электростанции\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                       SystemSounds.Beep.Play();
                       return;
                   }
                
                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where STATION_NAME=" + "'" + comboBox1.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
                else
                {

                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " and STATION_NAME=" + "'" + comboBox1.Text + "' or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
            }
            //////////////////////////////

            //////////////////////////////
            if (checkBox2.Checked == true)
            {
                if (comboBox2.Text.Length == 0)
                {
                    MessageBox.Show("Не заполнено поле \"Наименование оборудования\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    SystemSounds.Beep.Play();
                    return;
                }

                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where EQUIPMENT_NAME=" + "'" + comboBox2.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " and EQUIPMENT_NAME=" + "'" + comboBox2.Text + "' or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
            }
            //////////////////////////////

            //////////////////////////////
            if (checkBox3.Checked == true)
            {
                if (comboBox3.Text.Length == 0)
                {
                    MessageBox.Show("Не заполнено поле \"Вид работ\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    SystemSounds.Beep.Play();
                    return;
                }

         
                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where TYPE_WORK=" + "'" + comboBox3.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " and TYPE_WORK=" + "'" + comboBox3.Text + "' or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
            }
            ////////////////////////////////

            //////////////////////////////
            if (checkBox4.Checked == true)
            {
                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where PERIOD_REPAIR_BEGIN>=convert(datetime,'" + dateTimePicker1.Value.ToShortDateString() + " 00:00:00.000', 103) and PERIOD_REPAIR_END<=convert(datetime,'" + dateTimePicker2.Value.ToShortDateString() + " 23:59:59.000', 103)", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " and PERIOD_REPAIR_BEGIN>=convert(datetime,'" + dateTimePicker1.Value.ToShortDateString() + " 00:00:00.000', 103) and PERIOD_REPAIR_END<=convert(datetime,'" + dateTimePicker2.Value.ToShortDateString() + " 23:59:59.000', 103) or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
            }
            ////////////////////////////////

            //////////////////////////////
            if (checkBox5.Checked == true)
            {
                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }
            }
            ////////////////////////////////


        }

        private void Form3_Load(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch {}

            if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
            {
                SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];
                fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from ARCHIVE_REPAIRS_DATA where ID_USER=" + Form2.val + " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "ARCHIVE_REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];
                fill_gridview();
            }
            
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
        }

        
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox5.Checked = false;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
        }
    }
}
