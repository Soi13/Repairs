using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Data.SqlClient;
using System.IO;
using System.Media;

namespace Repairs
{
    public partial class Form5 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=REPAIRS;Data Source=T1212-W00079\MSSQLSERVER2012");
            
        public Image image1;

        public Form5()
        {
            InitializeComponent();
        }

        protected new void Resize(int width, int height, string adressin, string adressout)
        {
            Image img = Image.FromFile(adressin);
            Bitmap newImage;
            Graphics g;
            float kw;
            float kh;

            newImage = new Bitmap(width, height);
            g = Graphics.FromImage(newImage);
            kw = ((float)width) / ((float)img.Width);
            kh = ((float)height) / ((float)img.Height);
            g.Transform = new Matrix(kw, 0, 0, kh, 0, 0);
            g.DrawImageUnscaled(img, 0, 0);
            newImage.Save(adressout, ImageFormat.Jpeg);

        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Text = "Выбранный файл: "+openFileDialog1.FileName;
                Resize(400,300,openFileDialog1.FileName,Environment.CurrentDirectory+@"\img\1.jpg");
                image1 = Image.FromFile(Environment.CurrentDirectory + @"\img\1.jpg");
                pictureBox1.Image=image1;
             }
                      
                        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.FileName.Length==0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрано изображение.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
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
            Form1 form1 = new Form1();
            SqlCommand mycommand = new SqlCommand("insert into IMAGES_REPAIRS (IMAGE1,REPAIRS_ID) values (@img,"+Form1.grid_val.ToString()+")", conn);
            SqlParameter sqlParameter = new SqlParameter("@img", SqlDbType.VarBinary);
            string fileName = Environment.CurrentDirectory + @"\img\1.jpg";                  //Путь к файлу
            Image image = Image.FromFile(fileName);                                          //Изображение из файла.
            MemoryStream memoryStream = new MemoryStream();                                  //Поток в который запишем изображение
            image.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Bmp);
            sqlParameter.Value = memoryStream.ToArray();
            mycommand.Parameters.Add(sqlParameter);
            mycommand.ExecuteNonQuery();  
            conn.Close();
            memoryStream.Dispose();
            image.Dispose();
            image1.Dispose();
            File.Delete(Environment.CurrentDirectory + @"\img\1.jpg");

            SystemSounds.Beep.Play();
            MessageBox.Show("Изображение добавлено удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1();

            ////////Обновление отображения данных в гриде
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

            if ((Form2.val == 34) || (Form2.val == 31) || (Form2.val == 2))
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "REPAIRS_DATA");
                form1.dataGridView1.DataSource = ds.Tables[0];
                form1.fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from REPAIRS_DATA where ID_USER=" + Form2.val, conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "REPAIRS_DATA");
                form1.dataGridView1.DataSource = ds.Tables[0];
                form1.fill_gridview();
                /////////////////////////////////////////////
            }
            /////////////////////////////////////////////
        }
    }
}
