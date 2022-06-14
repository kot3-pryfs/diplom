using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Monitoring_performance
{
    public partial class AdminForm : Form
    {
        public string NameA = "";
        static string connectionString = @"Data Source=DESKTOP-R4EH7FQ; Initial Catalog = Мониторинг успеваемости; Integrated Security = True";
        public AdminForm()
        {
            InitializeComponent();
            MessageBox.Show("Первым выполняется перевод учащихся, затем копирование (Не перепутать последовательность!)", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AdminForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Date_Pass". При необходимости она может быть перемещена или удалена.
            this.date_PassTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Date_Pass);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Kyrs". При необходимости она может быть перемещена или удалена.
            this.kyrsTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Kyrs);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Specialnost". При необходимости она может быть перемещена или удалена.
            this.specialnostTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Specialnost);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Kyrator". При необходимости она может быть перемещена или удалена.
            this.kyratorTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Kyrator);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Prepod". При необходимости она может быть перемещена или удалена.
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Student". При необходимости она может быть перемещена или удалена.
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Gryppa". При необходимости она может быть перемещена или удалена.
            this.gryppaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Gryppa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);

        }

        private void скопироватьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool Buff = false;
            int God = Convert.ToInt32(DateTime.Today.Year);
            DateTime date = Convert.ToDateTime(this.date_PassTableAdapter.GetDataBy(God, true).Rows[0]["Date"]);
            if (date.Year == DateTime.Today.Year)
            {
                string message = "В текущем году данные уже копировались, вы действительно хотите скопировать их ещё раз?";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.YesNo;
                MessageBoxIcon icon = MessageBoxIcon.Warning;
                DialogResult resultt = MessageBox.Show(message, caption, button, icon);
                if (resultt == DialogResult.Yes)
                {
                    MessageBox.Show("Последнее копирование данных было выполнено " + date);
                    Buff = true;
                }
            }
            else
            {
                Buff = true;

            }

            if (Buff == true)
            {
                int result = 0;
                string sqlExpression = "Alter_Copy";
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand(sqlExpression, connection);
                        command.CommandType = CommandType.StoredProcedure;
                        result = command.ExecuteNonQuery();
                        MessageBox.Show("Данные скопированны. Обработанно " + result + " строк");
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка копирования данных. Проверьте подключение к сети или обратитесь к системному администратору(т.е. к себе:D)");
                }

                predmet_PrepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet_Prepod);
                gryppaTableAdapter.Fill(мониторинг_успеваемостиDataSet.Gryppa);
                studentTableAdapter.Fill(мониторинг_успеваемостиDataSet.Student);

                int count = dataGridView2.RowCount;
                for (int i = 0; i < count - 1; i++)
                {
                    int idGrypp = Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value); //извлекаем айдишник группы из датагрида
                    string NameGrypp = Convert.ToString(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["Nazvanie"]);    //извлекаем имя группы
                    int idSpec = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdSpecialnost"]);
                    gryppaTableAdapter.UpdateKyrsNameG(Convert.ToInt32(NameGrypp[1].ToString()), Convert.ToInt32(NameGrypp[2].ToString()), idGrypp);
                }

                predmet_PrepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet_Prepod);
                gryppaTableAdapter.Fill(мониторинг_успеваемостиDataSet.Gryppa);
                studentTableAdapter.Fill(мониторинг_успеваемостиDataSet.Student);
                this.date_PassTableAdapter.InsertQuery(DateTime.Today.ToString(), NameA, true);
            }
        }

        private void осуществитьПереводToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool Buff = false;
            int God = Convert.ToInt32(DateTime.Today.Year);
            DateTime date = Convert.ToDateTime(this.date_PassTableAdapter.GetDataBy(God, false).Rows[0]["Date"]);
            string NameAA = Convert.ToString(this.date_PassTableAdapter.GetDataBy(God, false).Rows[0]["Who_Pass"]);
            if (date.Year == DateTime.Today.Year)
            {
                string message = "В текущем году учащиеся уже переводились, вы действительно хотите перевести их ещё раз?";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.YesNo;
                MessageBoxIcon icon = MessageBoxIcon.Warning;
                DialogResult resultt = MessageBox.Show(message, caption, button, icon);
                if (resultt == DialogResult.Yes)
                {
                    MessageBox.Show("Последний перевод был выполнен " + date + " пользователем " + NameAA);
                    Buff = true;
                }
            }
            else
            {
                Buff = true;

            }

            if (Buff == true)
            {


                int result = 0;
                int count = dataGridView4.RowCount;

                string sqlExpression = "Pass";   //указываем название хранимой процедуры
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    command.CommandType = CommandType.StoredProcedure;
                    result = command.ExecuteNonQuery();
                    MessageBox.Show("Данные скопированны. Обработанно " + result + " строк");
                }

                for (int i = 0; i < count - 1; i++)
                {
                    int idSpecialnost = Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);  //извлекаем айдишник специальности из невидимого датагрида

                    int idGrypp = Convert.ToInt32(dataGridView4.Rows[i].Cells[0].Value); //извлекаем айдишник группы из датагрида

                    if (result != 0)
                    {
                        int IdSpec = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdSpecialnost"]);  // по аналогии извлекаем ID из других таблиц стандартным select
                        string NameG = Convert.ToString(specialnostTableAdapter.GetDataNameG(IdSpec).Rows[0]["NameG"]); // тут достаём букву группы
                        int IdKyrs = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdKyrs"]); // тут номер курса что по сути и есть ID курса
                        int NumberG = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["Number_Group"]); // тут номер группы
                        string NameGrypp = NameG + Convert.ToString(IdKyrs) + Convert.ToString(NumberG); //формируем название группы после копирования/перевода
                        gryppaTableAdapter.UpdateNameGrypp(NameGrypp, idGrypp, IdSpec);
                    }

                }
                predmet_PrepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet_Prepod);
                gryppaTableAdapter.Fill(мониторинг_успеваемостиDataSet.Gryppa);
                studentTableAdapter.Fill(мониторинг_успеваемостиDataSet.Student);
                this.date_PassTableAdapter.InsertQuery(DateTime.Today.ToString(), NameA, false);
            }
        }

        private void dataGridView4_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                int rowcount = dataGridView4.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView4.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса. Студентов такой группы не существует");
            }
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Student". При необходимости она может быть перемещена или удалена.
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
        }
    }
}