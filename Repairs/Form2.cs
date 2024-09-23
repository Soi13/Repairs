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
using System.Security.Cryptography; 

namespace Repairs
{
    public partial class Form2 : Form
    {
        public static string name_user;
        public static string login_user;
        public static int val;
        public static int button_period;
        public static int button_status;
        SqlDataAdapter da;
           

        public Form2()
        {
            InitializeComponent();
            
        }

        //Функция шифрования с помощью алгоритма MD5
        string GetHashString(string s)
        {
            //переводим строку в байт-массим  
            byte[] bytes = Encoding.Unicode.GetBytes(s);

            //создаем объект для получения средст шифрования  
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();

            //вычисляем хеш-представление в байтах  
            byte[] byteHash = CSP.ComputeHash(bytes);

            string hash = string.Empty;

            //формируем одну цельную строку из массива  
            foreach (byte b in byteHash)
                hash += string.Format("{0:x2}", b);

            return hash;
        }
        /////////////////////////////////////////

        private void button1_Click(object sender, EventArgs e)
        {

            if ((textBox1.Text.Length == 0) || (maskedTextBox1.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Заполнены не все поля!", "Внимание", MessageBoxButtons.OK);
                return;
            }

            string pass = GetHashString(maskedTextBox1.Text);

            SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=REPAIRS;Data Source=T1212-W00079\MSSQLSERVER2012");
                    
            conn.Open();
            SqlCommand mycommand = new SqlCommand("select * from users where user_name=" + "'" + textBox1.Text + "' and passw='" + pass + "'",conn);

            da = new SqlDataAdapter(mycommand);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "REPAIRS_DATA");
            if (ds.Tables[0].Rows.Count != 0)
            {
                object value = ds.Tables[0].Rows[0][0].ToString();
                if (value != null)
                {
                    Form1 form1 = new Form1();
                    val = Convert.ToInt16(value);
                    name_user = "Пользователь: " + textBox1.Text;
                    button_period = Convert.ToInt16(ds.Tables[0].Rows[0][4]);
                    button_status = Convert.ToInt16(ds.Tables[0].Rows[0][5]);
                    login_user = Convert.ToString(ds.Tables[0].Rows[0][1]);
                    form1.Show();
                    this.Visible = false;
                }

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не верно введено либо имя либо пароль!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        
            
                        
        }

        

       
    }
}
