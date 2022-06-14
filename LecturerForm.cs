using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Monitoring_performance
{
    public partial class LecturerForm : Form
    {
        string connectionString = @"Data Source=DESKTOP-R4EH7FQ; Initial Catalog = Мониторинг успеваемости; Integrated Security = True";
        public LecturerForm()
        {
            InitializeComponent();
        }

        private void LecturerForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Monitoring_View". При необходимости она может быть перемещена или удалена.
            this.monitoring_ViewTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Monitoring_View);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.God". При необходимости она может быть перемещена или удалена.
            this.godTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.God);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);
            //dataGridView2.DataSource = predmet_PrepodTableAdapter.GetData();
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Vid_Monitoringa". При необходимости она может быть перемещена или удалена.
            this.vid_MonitoringaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Vid_Monitoringa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Prepod". При необходимости она может быть перемещена или удалена.
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Gryppa". При необходимости она может быть перемещена или удалена.
            this.gryppaTableAdapter.FillBySecretar(this.мониторинг_успеваемостиDataSet.Gryppa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Student". При необходимости она может быть перемещена или удалена.
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring". При необходимости она может быть перемещена или удалена.
            this.monitoringTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Monitoring);

            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox1.Text = "";
            comboBox2.Text = "";
        }

      
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                string NameG = Convert.ToString(dataGridView1.Rows[rowcount].Cells[1].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod,idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса. Студентов такой группы не существует");
            }
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGryppa = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                int count = Convert.ToInt32(this.studentTableAdapter.ScalarQuery(idGryppa));
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                int idVid = Convert.ToInt32(this.vid_MonitoringaTableAdapter.GetDataBy(comboBox1.Text).Rows[0]["IdVidMonitoringa"]);
                int idGod = Convert.ToInt32(this.godTableAdapter.GetDataBy(comboBox2.Text).Rows[0]["IdGod"]);
                int result = 0;
                string sqlExpression = "ScalarQuery";
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand(sqlExpression, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        SqlParameter pGryppa = new SqlParameter
                        {
                            ParameterName = "@pGryppa",
                            Value = idGryppa
                        };
                        command.Parameters.Add(pGryppa);

                        SqlParameter pVid = new SqlParameter
                        {
                            ParameterName = "@pVid",
                            Value = idVid
                        };
                        command.Parameters.Add(pVid);

                        SqlParameter pGod = new SqlParameter
                        {
                            ParameterName = "@pGod",
                            Value = idGod
                        };
                        command.Parameters.Add(pGod);

                        SqlParameter pPred_prep = new SqlParameter
                        {
                            ParameterName = "@pPred_prep",
                            Value = idPred_Prep
                        };
                        command.Parameters.Add(pPred_prep);
                        result = Convert.ToInt32(command.ExecuteScalar());
                        MessageBox.Show("Обработанно " + result + " строк");
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка формирования мониторинга.");
                }
                int query = result;
                if (query == 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataBy2(idGryppa).Rows[i]["IdStudent"]);
                        this.monitoringTableAdapter.InsertQuery(0, idPred_Prep, idVid, idGod, idStudent);
                    }
                }
                else
                {
                    MessageBox.Show("Мониторинг данной группы уже сформирован", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGryppa);
                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
                comboBox1.Text = "";
                comboBox2.Text = "";
            }
            catch
            {
                string message = "Некорректный ввод данных или ошибка выполнения запроса. Попробуйте ещё раз";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Error;
                DialogResult result = MessageBox.Show(message, caption, button, icon);
            }
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod, idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                string NameG = Convert.ToString(dataGridView1.Rows[rowcount].Cells[1].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod, idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса");
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                string NameG = Convert.ToString(dataGridView1.Rows[rowcount].Cells[1].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod, idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                string NameG = Convert.ToString(dataGridView1.Rows[rowcount].Cells[1].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod, idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса");
            }
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                string NameG = Convert.ToString(dataGridView1.Rows[rowcount].Cells[1].Value);
                rowcount = dataGridView2.CurrentCell.RowIndex;
                int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                this.studentTableAdapter.FillBySelect(мониторинг_успеваемостиDataSet.Student, idGroup);
                this.predmet_PrepodTableAdapter.FillByLT(мониторинг_успеваемостиDataSet.Predmet_Prepod, idGroup);
                this.monitoring_ViewTableAdapter.FillBySearch(мониторинг_успеваемостиDataSet.Monitoring_View, idPred_Prep, comboBox1.Text, comboBox2.Text, idGroup);
            }
            catch
            {
                MessageBox.Show("Ошибка выполнения запроса");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int rowcount = dataGridView1.CurrentCell.RowIndex;
            int idGryppa = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
            int count = Convert.ToInt32(this.studentTableAdapter.ScalarQuery(idGryppa));
            rowcount = dataGridView2.CurrentCell.RowIndex;
            int idPred_Prep = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
            for (int i = 0; i < count; i++)
            {
                int Ocenka = Convert.ToInt32(dataGridView3.Rows[i].Cells[2].Value);
                int idMonit = Convert.ToInt32(dataGridView3.Rows[i].Cells[1].Value);
                this.monitoringTableAdapter.UpdateQuery(Ocenka, idMonit);
            }
        }
    }
}
