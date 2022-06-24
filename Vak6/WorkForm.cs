using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Vak6
{
    enum RowState
    {
        Existed,
        New,
        Modified,
        ModifiedNew,
        Deleted
    }
    public partial class WorkForm : Form
    {
        private SqlConnection SqlConnection = null;

        private SqlCommandBuilder SqlBuilder = null;

        private SqlDataAdapter SqlDataAdapter = null;

        private DataSet DataSet = null;

        private bool newRowAdding = false;

        DataBase database = new DataBase();

        int selectedRow;
        public WorkForm()
        {
            InitializeComponent();
        }

        ////Создание столбцов
        //private void CreateColums()
        //{

        //    dataGridView1.Columns.Add("id", "Таб. номер");
        //    dataGridView1.Columns.Add("FIO", "ФИО");
        //    dataGridView1.Columns.Add("Podr", "Подразделение");
        //    dataGridView1.Columns.Add("Dolj", "Должность");
        //    dataGridView1.Columns.Add("Date_born", "Дата рождения");
        //    dataGridView1.Columns.Add("Date1", "Дата 1 прививки");
        //    dataGridView1.Columns.Add("Data2_Layt", "Дата 2 прививки или Лайт");
        //    dataGridView1.Columns.Add("NumbS", "Номер сертификата");
        //    dataGridView1.Columns.Add("Revak1k", "Ревакцинация 1К");
        //    dataGridView1.Columns.Add("Revak2k_Layt", "Ревакцинация 2К или Лайт");
        //    dataGridView1.Columns.Add("Data_Prikaz", "Дата ознакомления с приказом");
        //    dataGridView1.Columns.Add("Data_Otkaz", "Дата отказа от вакцинации");
        //    dataGridView1.Columns.Add("Data_Otstran", "Дата отстранения от работы");
        //    dataGridView1.Columns.Add("Data_Dopysk", "Дата допуска к работе");
        //    dataGridView1.Columns.Add("IsNew", String.Empty);
        //    dataGridView1.Columns[4].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[5].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[6].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[8].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[9].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[10].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[11].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[12].DefaultCellStyle.Format = "dd.mm.yyyy";
        //    dataGridView1.Columns[13].DefaultCellStyle.Format = "dd.mm.yyyy";
        //}

        ////Типы даннх стоблцов
        //private void ReadSingleRow(DataGridView dgw, IDataRecord record)
        //{
        //    dgw.Rows.Add(record.GetInt32(0), record.GetString(1), record.GetString(2), record.GetString(3), record.GetString(4), record.GetString(5), record.GetString(6), record.GetInt32(7), record.GetString(8), record.GetString(9), record.GetString(10), record.GetString(11), record.GetString(12), record.GetString(13), RowState.ModifiedNew);

        //}


        ////Вывод данных из таблицы
        //private void RefreshDataGrid(DataGridView dgw)
        //{
        //    dgw.Rows.Clear();

        //    string queryString = $"select * from Users";

        //    SqlCommand command = new SqlCommand(queryString, database.getConnection());

        //    database.openConnection();

        //    SqlDataReader reader = command.ExecuteReader();

        //    while (reader.Read())
        //    {
        //        ReadSingleRow(dgw, reader);
        //    }
        //    reader.Close();
        //}

        //private void WorkForm_Load(object sender, EventArgs e)
        //{
        //    CreateColums();
        //    RefreshDataGrid(dataGridView1);
        //}

        //private void Update1()
        //{
        //    for (int index = 0; index < dataGridView1.Rows.Count; index++)
        //    {
        //        var rowState = (RowState)dataGridView1.Rows[index].Cells[0].Value;

        //        if (rowState == RowState.Existed)
        //            continue;

        //        if (rowState == RowState.Modified)
        //        {
        //            var id = dataGridView1.Rows[index].Cells[0].Value.ToString();
        //            var Data1p = dataGridView1.Rows[index].Cells[5].Value.ToString();
        //            var Data3p = dataGridView1.Rows[index].Cells[6].Value.ToString();
        //            var NumS = dataGridView1.Rows[index].Cells[7].Value.ToString();
        //            var Rev1k = dataGridView1.Rows[index].Cells[8].Value.ToString();
        //            var Rev2k = dataGridView1.Rows[index].Cells[9].Value.ToString();
        //            var DataP = dataGridView1.Rows[index].Cells[10].Value.ToString();
        //            var DataO = dataGridView1.Rows[index].Cells[11].Value.ToString();
        //            var DataOt = dataGridView1.Rows[index].Cells[12].Value.ToString();
        //            var DataD = dataGridView1.Rows[index].Cells[13].Value.ToString();

        //            var changeQuery = $"update Users set Data1 = '{Data1p}',Data2_Layt = '{Data3p}',NumbS = '{NumS}',Revak1k = '{Rev1k}', Revak2k_Layt = '{Rev2k}', Data_Prikaz = '{DataP}', Data_Otkaz = '{DataO}', Data_Otstran = '{DataOt}', Data_Dopysk = '{DataD}' where id = '{id}'";

        //            var command = new SqlCommand(changeQuery, database.getConnection());
        //            command.ExecuteNonQuery();

        //        }
        //    }
        //}
        private void ExportToExcel_Click(object sender, EventArgs e)
        {

        }

        //Открытие инструкции
        private void инструкцияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            var document = new Microsoft.Office.Interop.Word.Document();
            document = application.Documents.Add(Template: @"C:\Users\Alexandr1\source\repos\Vak6\Vak6\Инструкция\Инструкция.docx");
            application.Visible = true;
            document.SaveAs(FileName: @"C:\Users\Alexandr1\source\repos\Vak6\Vak6\Инструкция\Инструкция.docx");
        }
        
        //Закрытие программы
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //private void buttom_update_Click(object sender, EventArgs e)
        //{
        //    //Сортировка по ФИО
        //    usersBindingSource.Filter = "[Podr] LIKE '" + textbox_search.Text + "%'";
        //}

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //Подставление данных в textbox
            selectedRow = e.RowIndex;


            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[selectedRow];

                //Data1.Text = row.Cells[5].Value.ToString();
                //Data3.Text = row.Cells[6].Value.ToString();
                //NumberS.Text = row.Cells[7].Value.ToString();
                //Revak1k.Text = row.Cells[8].Value.ToString();
                //Revak2k.Text = row.Cells[9].Value.ToString();
                //DataPrikaz.Text = row.Cells[10].Value.ToString();
                //DataOtkaz.Text = row.Cells[11].Value.ToString();
                //DataOtstranenie.Text = row.Cells[12].Value.ToString();
                //DataDopysk.Text = row.Cells[13].Value.ToString();
            }
        }


  
        //Сортировка по ФИО
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"ФИО LIKE '%{textbox_search.Text}%'";
        }

        //private void Change()
        //{
        //    var selectedRowIndex = dataGridView1.CurrentCell.RowIndex;

        //    var Data1p = Data1.Text;
        //    var Data3p = Data3.Text;
        //    var NumS = NumberS.Text;
        //    var Rev1k = Revak1k.Text;
        //    var Rev2k = Revak2k.Text;
        //    var DataP = DataPrikaz.Text;
        //    var DataO = DataOtkaz.Text;
        //    var DataOt = DataOtstranenie.Text;
        //    var DataD = DataDopysk.Text;

        //    int NumbS;

        //    if (dataGridView1.Rows[selectedRowIndex].Cells[0].Value.ToString() != string.Empty)
        //    {
        //        if (int.TryParse(NumberS.Text, out NumbS))
        //        {
        //            dataGridView1.Rows[selectedRowIndex].SetValues(Data1p, Data3p, NumS, Rev1k, Rev2k, DataP, DataO, DataOt, DataD);
        //            dataGridView1.Rows[selectedRowIndex].Cells[13].Value = RowState.Modified;
        //        }
        //        else
        //        {
        //            MessageBox.Show("Неверно указаны данные.");
        //        }
        //    }
        //}
        private void button_change_Click(object sender, EventArgs e)
        {
            
        }

        private void Data1_TextChanged(object sender, EventArgs e)
        {

        }

        private void LoadData()
        {
            try
            {
                SqlDataAdapter = new SqlDataAdapter("SELECT *, 'Delete' AS [Команда] FROM Users", SqlConnection);

                SqlBuilder = new SqlCommandBuilder(SqlDataAdapter);

                SqlBuilder.GetInsertCommand();
                SqlBuilder.GetUpdateCommand();
                SqlBuilder.GetDeleteCommand();

                DataSet = new DataSet();

                SqlDataAdapter.Fill(DataSet, "Users");

                dataGridView1.DataSource = DataSet.Tables["Users"];

                for(int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[14,i] = linkCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Обновить данные в DataGridView1
        private void ReloadData()
        {
            try
            {
                DataSet.Tables["Users"].Clear();


                SqlDataAdapter.Fill(DataSet, "Users");

                dataGridView1.DataSource = DataSet.Tables["Users"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[14, i] = linkCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void WorkForm_Load(object sender, EventArgs e)
        {
            //Подключение к БД
            SqlConnection = new SqlConnection(@"Data Source=LAPTOP-D8FSJAKB;Initial Catalog=Users;Integrated Security=True");

            SqlConnection.Open();

            LoadData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if(e.ColumnIndex == 14)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();

                    if(task == "Delete")//Удаление данных из таблицы
                    {
                        if (MessageBox.Show("Удалить эту строку ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView1.Rows.RemoveAt(rowIndex);

                            DataSet.Tables["Users"].Rows[rowIndex].Delete();

                            SqlDataAdapter.Update(DataSet, "Users");
                        }
                    }
                    else if(task == "Insert")//Добавление данных в таблицу
                    { 
                        int rowIndex = dataGridView1.Rows.Count - 2;

                        DataRow row = DataSet.Tables["Users"].NewRow();

                        row["ФИО"] = dataGridView1.Rows[rowIndex].Cells["ФИО"].Value;
                        row["Подразделение"] = dataGridView1.Rows[rowIndex].Cells["Подразделение"].Value;
                        row["Должность"] = dataGridView1.Rows[rowIndex].Cells["Должность"].Value;
                        row["Дата рождения"] = dataGridView1.Rows[rowIndex].Cells["Дата рождения"].Value;
                        row["Дата 1 прививки"] = dataGridView1.Rows[rowIndex].Cells["Дата 1 прививки"].Value;
                        row["Дата 2 прививки или Лайт"] = dataGridView1.Rows[rowIndex].Cells["Дата 2 прививки или Лайт"].Value;
                        row["Номер сертификата"] = dataGridView1.Rows[rowIndex].Cells["Номер сертификата"].Value;
                        row["Ревакцинация 1К"] = dataGridView1.Rows[rowIndex].Cells["Ревакцинация 1К"].Value;
                        row["Ревакцинация 2К или Лайт"] = dataGridView1.Rows[rowIndex].Cells["Ревакцинация 2К или Лайт"].Value;
                        row["Дата ознокомления с приказом"] = dataGridView1.Rows[rowIndex].Cells["Дата ознокомления с приказом"].Value;
                        row["Дата отказа от вакцинации"] = dataGridView1.Rows[rowIndex].Cells["Дата отказа от вакцинации"].Value;
                        row["Дата отстранения от работы"] = dataGridView1.Rows[rowIndex].Cells["Дата отстранения от работы"].Value;
                        row["Дата допуска к работе"] = dataGridView1.Rows[rowIndex].Cells["Дата допуска к работе"].Value;

                        DataSet.Tables["Users"].Rows.Add(row);

                        DataSet.Tables["Users"].Rows.RemoveAt(DataSet.Tables["Users"].Rows.Count - 1);

                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);

                        dataGridView1.Rows[e.RowIndex].Cells[14].Value = "Delete";

                        SqlDataAdapter.Update(DataSet, "Users");

                        newRowAdding = false;
                    }
                    else if(task == "Update")//Редактирование и сохраниение данных в таблицу
                    {
                        int r = e.RowIndex;

                        DataSet.Tables["Users"].Rows[r]["ФИО"] = dataGridView1.Rows[r].Cells["ФИО"].Value;
                        DataSet.Tables["Users"].Rows[r]["Подразделение"] = dataGridView1.Rows[r].Cells["Подразделение"].Value;
                        DataSet.Tables["Users"].Rows[r]["Должность"] = dataGridView1.Rows[r].Cells["Должность"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата рождения"] = dataGridView1.Rows[r].Cells["Дата рождения"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата 1 прививки"] = dataGridView1.Rows[r].Cells["Дата 1 прививки"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата 2 прививки или Лайт"] = dataGridView1.Rows[r].Cells["Дата 2 прививки или Лайт"].Value;
                        DataSet.Tables["Users"].Rows[r]["Номер сертификата"] = dataGridView1.Rows[r].Cells["Номер сертификата"].Value;
                        DataSet.Tables["Users"].Rows[r]["Ревакцинация 1К"] = dataGridView1.Rows[r].Cells["Ревакцинация 1К"].Value;
                        DataSet.Tables["Users"].Rows[r]["Ревакцинация 2К или Лайт"] = dataGridView1.Rows[r].Cells["Ревакцинация 2К или Лайт"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата ознокомления с приказом"] = dataGridView1.Rows[r].Cells["Дата ознокомления с приказом"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата отказа от вакцинации"] = dataGridView1.Rows[r].Cells["Дата отказа от вакцинации"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата отстранения от работы"] = dataGridView1.Rows[r].Cells["Дата отстранения от работы"].Value;
                        DataSet.Tables["Users"].Rows[r]["Дата допуска к работе"] = dataGridView1.Rows[r].Cells["Дата допуска к работе"].Value;

                        SqlDataAdapter.Update(DataSet, "Users");

                        dataGridView1.Rows[e.RowIndex].Cells[14].Value = "Delete";

                    }
                    ReloadData();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void buttom_update_Click_1(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false)//При записи данных в пустую строку появляется кнопка Insert
                {
                    newRowAdding = true;

                    int lastRow = dataGridView1.RowCount - 2;

                    DataGridViewRow row = dataGridView1.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[14, lastRow] = linkCell;

                    row.Cells["Команда"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding == false)//При измениении данных в строке, вместо кнопки Delete появляется Update
                {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[14, rowIndex] = linkCell;

                    editingRow.Cells["Команда"].Value = "Update";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReloadData();
        }
        ////Ограничение на ввод символов
        //private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        //{
        //    e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);

        //    if (dataGridView1.CurrentCell.ColumnIndex == 4)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if(textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 5)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 6)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 7)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 8)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 9)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 10)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 11)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 12)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
        //    if (dataGridView1.CurrentCell.ColumnIndex == 13)
        //    {
        //        TextBox textbox = e.Control as TextBox;

        //        if (textbox != null)
        //        {
        //            textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
        //        }
        //    }
           
        //}

        //private void Column_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if(!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
        //    {
        //        e.Handled = true;
        //    }
        //}

        private void вгрузкаВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application xcelApp = new Excel.Application();

            xcelApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)xcelApp.ActiveSheet;
            int i, j;
                for (i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count-1; j++)
                    {
                    wsh.Cells[1, j+1] = dataGridView1.Columns[j].HeaderText;
                    wsh.Cells[i + 2, j + 1] = dataGridView1[j,i].Value.ToString();
                    }
                }
             
                xcelApp.Visible = true;
            
        }
    }
}
