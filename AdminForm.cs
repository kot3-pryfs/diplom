using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Monitoring_performance
{
    public partial class AdminForm : Form
    {
        static string connectionString = @"Data Source=DESKTOP-R4EH7FQ; Initial Catalog = Мониторинг успеваемости; Integrated Security = True";
        public AdminForm()
        {
            InitializeComponent();
        }

        private void AdminForm_Load(object sender, EventArgs e)
        {
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
                string NameGrypp = Convert.ToString(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["Nazvanie"]);    //извлекаем имя группы из датагрида
                int idSpec = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdSpecialnost"]);
                gryppaTableAdapter.UpdateKyrsNameG(Convert.ToInt32(Convert.ToString(NameGrypp[1])), Convert.ToInt32(Convert.ToString(NameGrypp[2])), idGrypp);
    
            }

            predmet_PrepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet_Prepod);
            gryppaTableAdapter.Fill(мониторинг_успеваемостиDataSet.Gryppa);
            studentTableAdapter.Fill(мониторинг_успеваемостиDataSet.Student);
        }

        private void осуществитьПереводToolStripMenuItem_Click(object sender, EventArgs e)
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
                    int IdSpec = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdSpecialnost"]);
                    string NameG = Convert.ToString(specialnostTableAdapter.GetDataNameG(IdSpec).Rows[0]["NameG"]);
                    int IdKyrs = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["IdKyrs"]);
                    int NumberG = Convert.ToInt32(gryppaTableAdapter.GetDataNameG(idGrypp).Rows[0]["Number_Group"]);
                    string NameGrypp = NameG + Convert.ToString(IdKyrs) + Convert.ToString(NumberG);
                    gryppaTableAdapter.UpdateNameGrypp(NameGrypp, idGrypp, IdSpec);
                }

            }
            predmet_PrepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet_Prepod);
            gryppaTableAdapter.Fill(мониторинг_успеваемостиDataSet.Gryppa);
            studentTableAdapter.Fill(мониторинг_успеваемостиDataSet.Student);
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