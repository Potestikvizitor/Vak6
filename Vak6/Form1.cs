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

namespace Vak6
{
    public partial class Form1 : Form
    {

        DataBase database = new DataBase();
        public Form1()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            //Регистрация пользователя(ввод логина и пароля)
            var loginUser = textbox_login.Text;
            var passUser = textbox_pass.Text;

            SqlDataAdapter adapter = new SqlDataAdapter();
            DataTable table = new DataTable();

            string querysring = $"select id_user, login_user, password_user from register where login_user = '{loginUser}' and password_user = '{passUser}'";

            SqlCommand command = new SqlCommand(querysring, database.getConnection());

            adapter.SelectCommand = command;
            adapter.Fill(table);

            if(table.Rows.Count == 1)
            {
                WorkForm workForm = new WorkForm();
                this.Hide();
                workForm.ShowDialog();
                this.Close();
            }
            else
            {
                MessageBox.Show("Неверноый логин или пароль");
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {   
            //Макс. кол-во симовлов в полях "Логин" и "Пароль"
            textbox_login.MaxLength = 50;
            textbox_pass.MaxLength = 50;
        }
    }
}
