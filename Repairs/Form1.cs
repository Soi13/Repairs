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
using System.Data.Common;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Repairs
{
    public partial class Form1 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=REPAIRS;Data Source=T1212-W00079\MSSQLSERVER2012");
                    
        SqlDataAdapter da;
        DataSet ds;
        Form2 form2 = new Form2();
        Form4 form4 = new Form4();
        Form5 form5 = new Form5();
        Form6 form6 = new Form6();

        public static int grid_val;
        public Image image1;

        public Form1()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["STATION_NAME"].HeaderText = "Наименование станции";
            dataGridView1.Columns["STATION_NAME"].Width = 100;
            dataGridView1.Columns["EQUIPMENT_NAME"].HeaderText = "Наименование оборудования";
            dataGridView1.Columns["EQUIPMENT_NAME"].Width = 100;
            dataGridView1.Columns["TYPE_WORK"].HeaderText = "Вид работ";
            dataGridView1.Columns["TYPE_WORK"].Width = 200;
            dataGridView1.Columns["TYPE_REPAIR"].HeaderText = "Вид ремонта";
            dataGridView1.Columns["TYPE_REPAIR"].Width = 70;
            dataGridView1.Columns["PERIOD_REPAIR_BEGIN"].HeaderText = "Начало ремонта";
            dataGridView1.Columns["PERIOD_REPAIR_BEGIN"].Width = 100;
            dataGridView1.Columns["PERIOD_REPAIR_END"].HeaderText = "Окончание ремонта";
            dataGridView1.Columns["PERIOD_REPAIR_END"].Width = 100;
            dataGridView1.Columns["MOVE_PERIOD_REPAIR"].HeaderText = "Перенос срока окончания";
            dataGridView1.Columns["MOVE_PERIOD_REPAIR"].Width = 100;
            dataGridView1.Columns["EXPLAIN_MOVE_REPAIR"].HeaderText = "Обоснование переноса срока";
            dataGridView1.Columns["EXPLAIN_MOVE_REPAIR"].Width = 250;
            dataGridView1.Columns["VOLUME"].HeaderText = "Плановый объем";
            dataGridView1.Columns["VOLUME"].Width = 150;
            dataGridView1.Columns["VIPOLNENO_TEK_DATA"].HeaderText = "Выполнено на тек. дату";
            dataGridView1.Columns["VIPOLNENO_TEK_DATA"].Width = 150;
            dataGridView1.Columns["PROCENT_VIPOLNEN"].HeaderText = "Процент выполнения";
            dataGridView1.Columns["PROCENT_VIPOLNEN"].Width = 100;
            dataGridView1.Columns["ID_USER"].Visible = false;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "Дата создания";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 100;
            dataGridView1.Columns["PERIOD"].Visible = false;
            dataGridView1.Columns["STATUS_UTVERJDEN"].Visible = false;
        }
        //////////////////////////////////////////////////////

                  
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                da.Update((System.Data.DataTable)dataGridView1.DataSource);
                button1.Enabled = true;
                button1.Text = "Редактирование";
                button2.Visible = false;
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в базе данных выполнены!","Уведомление о результатах", MessageBoxButtons.OK);

            }
            catch (Exception)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в базе данных выполнить не удалось!","Уведомление о результатах", MessageBoxButtons.OK);
            }

        }

      

        private void button4_Click(object sender, EventArgs e)
        {
            if ((comboBox1.Text.Length == 0) || (comboBox2.Text.Length == 0) || (comboBox3.Text.Length == 0) || (comboBox4.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не все поля заполнены.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            //Чтение текущего периода из таблицы
            conn.Open();
            SqlCommand mycommand = new SqlCommand("select * from PERIOD", conn);
            object value = mycommand.ExecuteScalar();
            conn.Close();
            ////////////////////////////////////

            /////////Вставка данных
            SqlCommand scmd = conn.CreateCommand();
            scmd.CommandText = "INSERT INTO REPAIRS_DATA (STATION_NAME,EQUIPMENT_NAME,TYPE_WORK,TYPE_REPAIR,PERIOD_REPAIR_BEGIN,PERIOD_REPAIR_END,MOVE_PERIOD_REPAIR,EXPLAIN_MOVE_REPAIR,VOLUME,VIPOLNENO_TEK_DATA,PROCENT_VIPOLNEN,ID_USER,DATETIME_CREATE,PERIOD,STATUS_UTVERJDEN ) VALUES (" + "'" + comboBox1.Text + "', '" + comboBox2.Text + "', '" + comboBox3.Text + "', '" + comboBox4.Text + "', convert(datetime,'" + dateTimePicker1.Value + "', 103), convert(datetime,'" + dateTimePicker2.Value + "', 103), convert(datetime,'" + dateTimePicker3.Value + "', 103), '" + richTextBox1.Text + "', '" + textBox1.Text + "', '" + textBox2.Text + "', '" + textBox3.Text + "%" + "', '" + Form2.val + "', convert(datetime,'" + DateTime.Now.ToString() + "', 103), '" + value.ToString() + "', '" + 0 + "')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader;
            reader = scmd.ExecuteReader();
            conn.Close();
            //////////////////

            ////////Обновление отображения данных в гриде
            try
            {
                conn.Open();
            }
            catch{ }


            if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];
                fill_gridview();
                ///////////////////////////////////////////// 
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA where ID_USER=" + Form2.val, conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];
                fill_gridview();
                ///////////////////////////////////////////// 
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {                                 
            //Включение отключение кнопок закрытия периода и подтверждения в зависимости от прав пользователя
            if (Form2.button_period == 1) { button5.Visible = true; } else { button5.Visible = false;}
            if (Form2.button_status == 1) { button3.Visible = true; button6.Visible = true; } else { button3.Visible = false; button6.Visible = false; }

            try
            {
                conn.Open();
            }
            catch{}

            if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];

                fill_gridview();
            }
            else
            {
                //SqlCommand command = new SqlCommand("select * from REPAIRS_DATA where ID_USER="+Form2.val, conn);
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA where ID_USER=" + Form2.val + " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "REPAIRS_DATA");
                dataGridView1.DataSource = ds.Tables[0];

                fill_gridview();
            }
            //Чтение текущего периода из таблицы
            conn.Open();
            SqlCommand mycommand = new SqlCommand("select * from PERIOD", conn);
            object value = mycommand.ExecuteScalar();
            conn.Close();
            ////////////////////////////////////
            label3.Text="Текущий период: " + value.ToString();
            label2.Text = Form2.name_user; 

            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую

            if (ds.Tables[0].Rows.Count != 0)
            {
                if (Convert.ToInt16(ds.Tables[0].Rows[0][15]) == 1)
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    button6.Enabled = true;////Данная кнопка видна только кураторам и тому кто закрывает период, т.е. Дашкину

                    contextMenuStrip1.Enabled = false;

                    SystemSounds.Beep.Play();
                    MessageBox.Show("Редактирование данных запрещено, т.к. ваш куратор утвердил данные ремонты.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    button6.Enabled = false;////Данная кнопка видна только кураторам и тому кто закрывает период, т.е. Дашкину
                    contextMenuStrip1.Enabled = true;
                }
            }

            /////Вставка данных в таблицу журнала вход/выход
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into JOURNAL (ID_USER,FIO,EVENT_DATETIME,EVENT_STATUS) VALUES (" + "'" + Form2.val + "', (select USER_NAME from USERS where ID=" + Form2.val + "), convert(datetime,'" + DateTime.Now.ToString() + "', 103),'Вход')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.Text = "Включен режим редактирования";
            button2.Visible = true;
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            button1.Enabled = false;
        }

       
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

          
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void архивыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        private void удалитьТекущуюЗаписьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить данную запись?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////////Удаление данных
                SqlCommand scmd = conn.CreateCommand();
                scmd.CommandText = "delete from REPAIRS_DATA where ID=" + dataGridView1.CurrentRow.Cells[0].Value;
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader;
                reader = scmd.ExecuteReader();
                conn.Close();
                //////////////////

                ////////Обновление отображения данных в гриде
                try
                {
                    conn.Open();
                }
                catch { }

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command = new SqlCommand("select * from REPAIRS_DATA", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                    /////////////////////////////////////////////
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from REPAIRS_DATA where ID_USER=" + Form2.val + " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "REPAIRS_DATA");
                    dataGridView1.DataSource = ds.Tables[0];
                    fill_gridview();
                }               
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите закрыть текущий период?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "REPAIRS_DATA");
                
                for (int i = 0; i <= ds.Tables[0].Rows.Count-1; i++)
                {
                    string ID = Convert.ToString(ds.Tables[0].Rows[i][0]);
                    string STATION_NAME = Convert.ToString(ds.Tables[0].Rows[i][1]);
                    string EQUIPMENT_NAME = Convert.ToString(ds.Tables[0].Rows[i][2]);
                    string TYPE_WORK = Convert.ToString(ds.Tables[0].Rows[i][3]);
                    string TYPE_REPAIR = Convert.ToString(ds.Tables[0].Rows[i][4]);
                    string PERIOD_REPAIR_BEGIN = Convert.ToString(ds.Tables[0].Rows[i][5]);
                    string PERIOD_REPAIR_END = Convert.ToString(ds.Tables[0].Rows[i][6]);
                    string MOVE_PERIOD_REPAIR = Convert.ToString(ds.Tables[0].Rows[i][7]);
                    string EXPLAIN_MOVE_REPAIR = Convert.ToString(ds.Tables[0].Rows[i][8]);
                    string VOLUME = Convert.ToString(ds.Tables[0].Rows[i][9]);
                    string VIPOLNENO_TEK_DATA = Convert.ToString(ds.Tables[0].Rows[i][10]);
                    string PROCENT_VIPOLNEN = Convert.ToString(ds.Tables[0].Rows[i][11]);
                    string ID_USER = Convert.ToString(ds.Tables[0].Rows[i][12]);
                    string DATETIME_CREATE = Convert.ToString(ds.Tables[0].Rows[i][13]);
                    string PERIOD = Convert.ToString(ds.Tables[0].Rows[i][14]);
                    string STATUS_UTVERJDEN = Convert.ToString(ds.Tables[0].Rows[i][15]);

                    if (Convert.ToDateTime(PERIOD_REPAIR_END) <= DateTime.Now)
                    {
                        /////////Вставка данных
                        SqlCommand scmd1 = conn.CreateCommand();
                        scmd1.CommandText = "INSERT INTO ARCHIVE_REPAIRS_DATA (STATION_NAME,EQUIPMENT_NAME,TYPE_WORK,TYPE_REPAIR,PERIOD_REPAIR_BEGIN,PERIOD_REPAIR_END,MOVE_PERIOD_REPAIR,EXPLAIN_MOVE_REPAIR,VOLUME,VIPOLNENO_TEK_DATA,PROCENT_VIPOLNEN,ID_USER,DATETIME_CREATE,PERIOD,STATUS_UTVERJDEN) " +
                                           "VALUES ('" + STATION_NAME + "','" + EQUIPMENT_NAME + "','" + TYPE_WORK + "','" + TYPE_REPAIR + "', convert(datetime,'" + PERIOD_REPAIR_BEGIN + "', 103), convert(datetime,'" + PERIOD_REPAIR_END + "', 103), convert(datetime,'" + MOVE_PERIOD_REPAIR + "', 103) ,'" + EXPLAIN_MOVE_REPAIR + "','" + VOLUME + "','" + VIPOLNENO_TEK_DATA + "','" + PROCENT_VIPOLNEN + "','" + ID_USER + "',convert(datetime,'" + DATETIME_CREATE + "', 103),'" + PERIOD + "','" + STATUS_UTVERJDEN + "')";
                        try
                        {
                            conn.Open();
                        }
                        catch {}
                        SqlDataReader reader1;
                        reader1 = scmd1.ExecuteReader();
                        conn.Close();
                        //////////////////

                        /////////Удаление данных из таблицы после переноса данных в архив
                        SqlCommand scmd3 = conn.CreateCommand();
                        scmd3.CommandText = "delete from REPAIRS_DATA where ID="+ID;
                        try
                        {
                            conn.Open();
                        }
                        catch {}
                        SqlDataReader reader3;
                        reader3 = scmd3.ExecuteReader();
                        conn.Close();
                        //////////////////
                    }
                    

                }
                
                //Чтение текущего периода из таблицы
                conn.Open();
                SqlCommand mycommand = new SqlCommand("select * from PERIOD", conn);
                object value = mycommand.ExecuteScalar();
                conn.Close();
                ////////////////////////////////////

                //Чтение текущего периода и увеличение текущей недели на 1
                string new_period="";
                string d = value.ToString().Substring(0, 1);
                if (Convert.ToInt32(d) == 4) { new_period = "1-я неделя/" + DateTime.Now.Month + "." + DateTime.Now.Year; }
                if (Convert.ToInt32(d) == 3) { new_period = "4-я неделя/" + DateTime.Now.Month + "." + DateTime.Now.Year; }
                if (Convert.ToInt32(d) == 2) { new_period = "3-я неделя/" + DateTime.Now.Month + "." + DateTime.Now.Year; }
                if (Convert.ToInt32(d) == 1) { new_period = "2-я неделя/" + DateTime.Now.Month + "." + DateTime.Now.Year; }
                ///////////////////////////////////

                /////////Обновление данных о периоде
                SqlCommand scmd2 = conn.CreateCommand();
                scmd2.CommandText = "UPDATE PERIOD SET PERIOD="+"'"+new_period+"'";
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader2;
                reader2 = scmd2.ExecuteReader();
                conn.Close();
                //////////////////

                //Чтение вновь открытого периода из таблицы
                conn.Open();
                SqlCommand mycommand1 = new SqlCommand("select * from PERIOD", conn);
                object value1= mycommand1.ExecuteScalar();
                conn.Close();
                ////////////////////////////////////

                ////////Обновление отображения данных в гриде
                try
                {
                    conn.Open();
                }
                catch {}

                if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
                {
                    SqlCommand command1 = new SqlCommand("select * from REPAIRS_DATA", conn);
                    da = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
                    DataSet ds1 = new DataSet();
                    conn.Close();
                    da.Fill(ds1, "REPAIRS_DATA");
                    dataGridView1.DataSource = ds1.Tables[0];
                    fill_gridview();
                }
                else
                {
                    SqlCommand command1 = new SqlCommand("select * from REPAIRS_DATA where ID_USER=" + Form2.val, conn);
                    da = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
                    DataSet ds1 = new DataSet();
                    conn.Close();
                    da.Fill(ds1, "REPAIRS_DATA");
                    dataGridView1.DataSource = ds1.Tables[0];
                    fill_gridview();
                    ///////////////////////////////////////////// 
                }
                
                label3.Text = "Текущий период: " + value1.ToString();
                SystemSounds.Beep.Play();
                MessageBox.Show("Новый период " + value1.ToString() + " открыт.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
            }
        }

        private void сменаПароляToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form4.ShowDialog();  
        }

       
        private void добавитьИзображенияКЗаписиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ////////////Проверка, существует ли привязанная картинка у записи, перед вставкой другой картриник
            try
            {
                conn.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с БД.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SystemSounds.Beep.Play();
                return;
            }
            SqlCommand mycommand = new SqlCommand("select * from IMAGES_REPAIRS where REPAIRS_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            SqlDataReader sqlDataReader = mycommand.ExecuteReader();
            if (sqlDataReader.HasRows)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данная запись уже имеет привязанное изображение. Необходимо удалить его, если хотите привязать другое.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                conn.Close();
                return;
            }
            ////////////////////////////////////////////////////////////////////////////
            conn.Close();
            grid_val = (int)dataGridView1.CurrentRow.Cells[0].Value;
            form5.ShowDialog();
        }

                
        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox3.Text)>100)  
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Процент выполнения больше 100 не может быть!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox3.Clear();
            }
        }
       

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch {}
            SqlCommand mycommand = new SqlCommand("select * from IMAGES_REPAIRS where REPAIRS_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            SqlDataReader sqlDataReader = mycommand.ExecuteReader();
            if (sqlDataReader.HasRows)
            {
                MemoryStream memoryStream = new MemoryStream();
                foreach (DbDataRecord record in sqlDataReader)
                    memoryStream.Write((byte[])record["IMAGE1"], 0, ((byte[])record["IMAGE1"]).Length);
                Image image = Image.FromStream(memoryStream);
                //image.Save(Environment.CurrentDirectory + @"\img\view.jpg");
                pictureBox1.Image = image;
                memoryStream.Dispose();
                //image.Dispose();
            }
            else
            {
                pictureBox1.Image = null;
            }
            conn.Close();
                  
        }

        private void удалитьИзображениеТекущейЗаписиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ////////////Проверка, существует ли привязанная картинка у записи, перед вставкой другой картриник
            try
            {
                conn.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с БД.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SystemSounds.Beep.Play();
                return;
            }
            SqlCommand mycommand = new SqlCommand("select * from IMAGES_REPAIRS where REPAIRS_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            SqlDataReader sqlDataReader = mycommand.ExecuteReader();
            if (!sqlDataReader.HasRows)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Удалять нечего! К записи изображение не привязано.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                conn.Close();
                return;
            }
            conn.Close();
            ////////////////////////////////////////////////////////////////////////////
            
        
            if (MessageBox.Show("Вы уверены, что хотите удалить изображение привязанное к текущей записи?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////////Удаление у текущей записи привязанного изображения
                SqlCommand scmd = conn.CreateCommand();
                scmd.CommandText = "delete from IMAGES_REPAIRS where REPAIRS_ID=" + dataGridView1.CurrentRow.Cells[0].Value;
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader;
                reader = scmd.ExecuteReader();
                conn.Close();
                //////////////////

                SystemSounds.Beep.Play();
                MessageBox.Show("Изображение удалено удачно!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);            
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите утвердить введеную информацию по ремонтам?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////////Подтверждение куратором введенных данных его подопечных
                SqlCommand scmd2 = conn.CreateCommand();
                scmd2.CommandText = "UPDATE REPAIRS_DATA SET STATUS_UTVERJDEN=1 where ID_USER=" + Form2.val+ " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER="+ Form2.val+")";
                try
                {
                    conn.Open();
                }
                catch {}
                SqlDataReader reader2;
                reader2 = scmd2.ExecuteReader();
                conn.Close();
                //////////////////  

                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
                button3.Enabled = false;
                button6.Enabled = true;
               
                SystemSounds.Beep.Play();
                MessageBox.Show("Утверждено! Редактирование данных позиций для исполнителей, с данного момента, невозможно.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);            
         
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите отменить утверждение информации по ремонтам?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////////Подтверждение куратором введенных данных его подопечных
                SqlCommand scmd2 = conn.CreateCommand();
                scmd2.CommandText = "UPDATE REPAIRS_DATA SET STATUS_UTVERJDEN=0 where ID_USER=" + Form2.val + " or ID_USER in (select USERS.ID from USERS, CURATOR where USERS.CURATOR_ID=CURATOR.ID and CURATOR.ID_USER=" + Form2.val + ")";
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader2;
                reader2 = scmd2.ExecuteReader();
                conn.Close();
                //////////////////  

                button1.Enabled = true;
                button2.Enabled = true;
                button4.Enabled = true;
                button6.Enabled = true;
                button3.Enabled = false;

                SystemSounds.Beep.Play();
                MessageBox.Show("Утверждение снято. Редактирование данных позиций для исполнителей, с данного момента, возможно!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void отчетОВыполненииПоРемонтуЗиСToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch {}

            SqlCommand command = new SqlCommand("select * from REPAIRS_DATA ORDER BY STATION_NAME", conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "REPAIRS_DATA");
            //dataGridView1.DataSource = ds.Tables[0];
           


           Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
           ExcelApp.Application.Workbooks.Add(Type.Missing);
          
           //делаем временно неактивным документ
           ExcelApp.Interactive = false;
           ExcelApp.EnableEvents = false;

           /*ExcelApp.Columns[1].ColumnWidth = 23;
           ExcelApp.Columns[2].ColumnWidth = 29;
           ExcelApp.Columns[3].ColumnWidth = 25;
           ExcelApp.Columns[4].ColumnWidth = 15;
           ExcelApp.Columns[5].ColumnWidth = 18;
           ExcelApp.Columns[6].ColumnWidth = 19;
           ExcelApp.Columns[7].ColumnWidth = 25;
           ExcelApp.Columns[8].ColumnWidth = 28;
           ExcelApp.Columns[9].ColumnWidth = 20;
           ExcelApp.Columns[10].ColumnWidth = 27;
           ExcelApp.Columns[11].ColumnWidth = 21;*/

           //ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке

           
           ExcelApp.Cells[1, 1] = "Наименование станции";
           ExcelApp.Cells[1, 2] = "Наименование оборудования";
           ExcelApp.Cells[1, 3] = "Вид работ";
           ExcelApp.Cells[1, 4] = "Вид ремонта";
           ExcelApp.Cells[1, 5] = "Начало ремонта";
           ExcelApp.Cells[1, 6] = "Окончание ремонта";
           ExcelApp.Cells[1, 7] = "Перенос срока окончания";
           ExcelApp.Cells[1, 8] = "Обоснование срока переноса";
           ExcelApp.Cells[1, 9] = "Плановый объем";
           ExcelApp.Cells[1, 10] = "Выполнено на текущую дату";
           ExcelApp.Cells[1, 11] = "Процент выполнения";
           
           for (int i = 0; i <= ds.Tables[0].Rows.Count-1; i++)
           {
               ExcelApp.Cells[i+2, 1] = Convert.ToString(ds.Tables[0].Rows[i][1]);
               ExcelApp.Cells[i + 2, 2] = Convert.ToString(ds.Tables[0].Rows[i][2]);
               ExcelApp.Cells[i + 2, 3] = Convert.ToString(ds.Tables[0].Rows[i][3]);
               ExcelApp.Cells[i + 2, 4] = Convert.ToString(ds.Tables[0].Rows[i][4]);
               ExcelApp.Cells[i + 2, 5] = Convert.ToString(ds.Tables[0].Rows[i][5]);
               ExcelApp.Cells[i + 2, 6] = Convert.ToString(ds.Tables[0].Rows[i][6]);
               ExcelApp.Cells[i + 2, 7] = Convert.ToString(ds.Tables[0].Rows[i][7]);
               ExcelApp.Cells[i + 2, 8] = Convert.ToString(ds.Tables[0].Rows[i][8]);
               ExcelApp.Cells[i + 2, 9] = Convert.ToString(ds.Tables[0].Rows[i][9]);
               ExcelApp.Cells[i + 2, 10] = Convert.ToString(ds.Tables[0].Rows[i][10]);
               ExcelApp.Cells[i + 2, 11] = Convert.ToString(ds.Tables[0].Rows[i][11]);

           }
            
            //Показываем ексель
            ExcelApp.Visible = true;

            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
          }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            /////Вставка данных в таблицу журнала вход/выход
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into JOURNAL (ID_USER,FIO,EVENT_DATETIME,EVENT_STATUS) VALUES (" + "'" + Form2.val + "', (select USER_NAME from USERS where ID=" + Form2.val + "), convert(datetime,'" + DateTime.Now.ToString() + "', 103),'Выход')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////
        }

        private void задатьВопросРазработчикуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form6.ShowDialog();
        }

     
       
      

      

    }

}
