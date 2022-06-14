using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Monitoring_performance
{
    public partial class SecretaryForm : Form
    {
        public SecretaryForm()
        {
            InitializeComponent();
        }

        private void SecretaryForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.RedactST". При необходимости она может быть перемещена или удалена.
            this.redactSTTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.RedactST);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
            //this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Specialnost". При необходимости она может быть перемещена или удалена.
            this.specialnostTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Specialnost);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Prepod". При необходимости она может быть перемещена или удалена.
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Kyrs". При необходимости она может быть перемещена или удалена.
            this.kyrsTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Kyrs);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Student". При необходимости она может быть перемещена или удалена.
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);            
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Kyrator". При необходимости она может быть перемещена или удалена.
            this.kyratorTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Kyrator);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Specialnost". При необходимости она может быть перемещена или удалена.
            this.specialnostTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Specialnost);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Gryppa". При необходимости она может быть перемещена или удалена.
            this.gryppaTableAdapter.FillBySecretar(this.мониторинг_успеваемостиDataSet.Gryppa);
            // 
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
            //
            comboBox21.SelectedIndex = -1;
            comboBox20.SelectedIndex = -1;
            comboBox19.SelectedIndex = -1;
            comboBox18.SelectedIndex = -1;
            comboBox16.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;



        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.gryppaTableAdapter.FillBySecretar(this.мониторинг_успеваемостиDataSet.Gryppa);
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int rowcount = dataGridView1.CurrentCell.RowIndex;
            int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
            this.gryppaTableAdapter.FillBy(мониторинг_успеваемостиDataSet.Gryppa, idGroup);
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string message = "Вы уверены что хотите удалить данную запись?";
            string caption = "Мониторинг успеваемости";
            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            DialogResult result = MessageBox.Show(message, caption, button,icon);
            if(result == DialogResult.Yes)
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[0].Value);
                gryppaTableAdapter.DeleteQuery(idGroup);
                this.gryppaTableAdapter.FillBySecretar(this.мониторинг_успеваемостиDataSet.Gryppa);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int idspec = Convert.ToInt32(specialnostTableAdapter.GetDataByIG(comboBox2.Text).Rows[0]["IdSpecialnost"]);
                int idkyrat = Convert.ToInt32(kyratorTableAdapter.GetDataByIG(comboBox3.Text).Rows[0]["IdKyrator"]);
                int idkyrs = Convert.ToInt32(comboBox4.Text);
                this.gryppaTableAdapter.InsertQuery(comboBox1.Text, idspec, idkyrat, idkyrs, Convert.ToInt32(comboBox5.Text), 0);
                this.gryppaTableAdapter.FillBySecretar(this.мониторинг_успеваемостиDataSet.Gryppa);
            }
            catch
            {
                string message = "Некорректный ввод данных. Попробуйте ещё раз";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Error;
                DialogResult result = MessageBox.Show(message, caption, button, icon);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }


        private void button12_Click(object sender, EventArgs e)
        {
            this.prepodTableAdapter.Fill(мониторинг_успеваемостиDataSet.Prepod);
        }

        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            int rowcount = dataGridView3.CurrentCell.RowIndex;
            int idPrepod = Convert.ToInt32(dataGridView3.Rows[rowcount].Cells[0].Value);
            this.prepodTableAdapter.FillBy(мониторинг_успеваемостиDataSet.Prepod, idPrepod);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if ((comboBox11.Text.Length > 2) && (comboBox12.Text.Length > 2) && (comboBox13.Text.Length > 2) && (comboBox14.Text.Length > 2))
            {
                try
                {
                    this.prepodTableAdapter.InsertQuery(comboBox11.Text, comboBox12.Text, comboBox13.Text, comboBox14.Text);
                }
                catch
                {
                    string message = "Некорректный ввод данных. Попробуйте ещё раз";
                    string caption = "Мониторинг успеваемости";
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBoxIcon icon = MessageBoxIcon.Error;
                    DialogResult result = MessageBox.Show(message, caption, button, icon);
                }
            }
            else
            {
                string message = "Не заполнены обязательные поля. Попробуйте ещё раз";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Error;
                DialogResult result = MessageBox.Show(message, caption, button, icon);
            }
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Prepod);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string message = "Вы уверены что хотите удалить данную запись?";
            string caption = "Мониторинг успеваемости";
            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            DialogResult result = MessageBox.Show(message, caption, button, icon);
            if (result == DialogResult.Yes)
            {
                int rowcount = dataGridView3.CurrentCell.RowIndex;
                int idGroup = Convert.ToInt32(dataGridView3.Rows[rowcount].Cells[0].Value);
                prepodTableAdapter.DeleteQuery(idGroup);
                this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Prepod);
            }
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            int rowcount = dataGridView2.CurrentCell.RowIndex;
            int idStudent = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
            this.studentTableAdapter.FillBy(мониторинг_успеваемостиDataSet.Student, idStudent);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string message = "Вы уверены что хотите удалить данную запись?";
            string caption = "Мониторинг успеваемости";
            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            DialogResult result = MessageBox.Show(message, caption, button, icon);
            if (result == DialogResult.Yes)
            {
                int rowcount = dataGridView2.CurrentCell.RowIndex;
                int idStudent = Convert.ToInt32(dataGridView2.Rows[rowcount].Cells[0].Value);
                studentTableAdapter.DeleteQuery(idStudent);
                this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if ((comboBox6.Text.Length > 2) && (comboBox7.Text.Length > 2) && (comboBox8.Text.Length > 2) && (comboBox10.Text.Length > 1))
            {
                try
                {
                    int idGrupp = Convert.ToInt32(gryppaTableAdapter.GetDataBySG(comboBox10.Text).Rows[0]["IdGryppa"]);
                    this.studentTableAdapter.InsertQuery(comboBox6.Text, comboBox7.Text, comboBox8.Text, Convert.ToInt32(comboBox9.Text), idGrupp);
                }
                catch
                {
                    string message = "Некорректный ввод данных. Попробуйте ещё раз";
                    string caption = "Мониторинг успеваемости";
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBoxIcon icon = MessageBoxIcon.Error;
                    DialogResult result = MessageBox.Show(message, caption, button, icon);
                }
            }
            else
            {
                string message = "Не заполнены обязательные поля. Попробуйте ещё раз";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Error;
                DialogResult result = MessageBox.Show(message, caption, button, icon);
            }
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Student);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            this.predmetTableAdapter.Fill(мониторинг_успеваемостиDataSet.Predmet);
        }

        private void dataGridView4_DoubleClick(object sender, EventArgs e)
        {
            int rowcount = dataGridView4.CurrentCell.RowIndex;
            int idPredmet = Convert.ToInt32(dataGridView4.Rows[rowcount].Cells[0].Value);
            this.predmetTableAdapter.FillBy(мониторинг_успеваемостиDataSet.Predmet, idPredmet);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string message = "Вы уверены что хотите удалить данную запись?";
            string caption = "Мониторинг успеваемости";
            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            DialogResult result = MessageBox.Show(message, caption, button, icon);
            if (result == DialogResult.Yes)
            {
                int rowcount = dataGridView4.CurrentCell.RowIndex;
                int idpredmet = Convert.ToInt32(dataGridView4.Rows[rowcount].Cells[0].Value);
                predmetTableAdapter.DeleteQuery(idpredmet);
                this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if  (comboBox18.Text.Length > 2)
            {
                try
                { 
                    this.predmetTableAdapter.InsertQuery(comboBox18.Text, comboBox17.Text, "");
                }
                catch
                {
                    string message = "Некорректный ввод данных. Попробуйте ещё раз";
                    string caption = "Мониторинг успеваемости";
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBoxIcon icon = MessageBoxIcon.Error;
                    DialogResult result = MessageBox.Show(message, caption, button, icon);
                }
            }
            else
            {
                string message = "Не заполнены обязательные поля. Попробуйте ещё раз";
                string caption = "Мониторинг успеваемости";
                MessageBoxButtons button = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Error;
                DialogResult result = MessageBox.Show(message, caption, button, icon);
            }
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet);
        }

        private void dataGridView5_DoubleClick(object sender, EventArgs e)
        {
            int rowcount = dataGridView4.CurrentCell.RowIndex;
            int idPredmet = Convert.ToInt32(dataGridView4.Rows[rowcount].Cells[0].Value);
            this.predmetTableAdapter.FillBy(мониторинг_успеваемостиDataSet.Predmet, idPredmet);
        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
            }
            catch
            {
                // так лень писать обработчик. пожалуй забью хер
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            comboBox21.SelectedIndex = -1;
            comboBox20.SelectedIndex = -1;
            comboBox19.SelectedIndex = -1;
            comboBox16.SelectedIndex = -1;
            comboBox21.Text = "";
            comboBox20.Text = "";
            comboBox19.Text = "";
            comboBox16.Text = "";
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void comboBox21_TextUpdate(object sender, EventArgs e)
        {
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void comboBox20_TextUpdate(object sender, EventArgs e)
        {
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void comboBox16_TextUpdate(object sender, EventArgs e)
        {
            this.predmet_PrepodTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet_Prepod, comboBox21.Text, comboBox20.Text, comboBox16.Text);
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            comboBox21.SelectedIndex = -1;
            comboBox20.SelectedIndex = -1;
            comboBox19.SelectedIndex = -1;
            comboBox16.SelectedIndex = -1;
            comboBox21.Text = "";
            comboBox20.Text = "";
            comboBox19.Text = "";
            comboBox16.Text = "";
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string message = "Вы уверены что хотите удалить данную запись?";
            string caption = "Мониторинг успеваемости";
            MessageBoxButtons button = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            DialogResult result = MessageBox.Show(message, caption, button, icon);
            if (result == DialogResult.Yes)
            {
                int rowcount = dataGridView5.CurrentCell.RowIndex;
                int idpredmet_prepod = Convert.ToInt32(dataGridView5.Rows[rowcount].Cells[0].Value);
                this.predmet_PrepodTableAdapter.DeleteQuery(idpredmet_prepod);
                this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Predmet_Prepod);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void comboBox6_TextUpdate(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void comboBox7_TextUpdate(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void comboBox10_TextUpdate(object sender, EventArgs e)
        {
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox8.Text = "";
            comboBox10.Text = "";
            this.redactSTTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet1.RedactST, comboBox6.Text, comboBox7.Text, comboBox10.Text);
        }

        private void tabPage2_Enter(object sender, EventArgs e)
        {
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox8.Text = "";
            comboBox10.Text = "";
        }

        private void импртироватьДанныеВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportExcel form = new ImportExcel();
            form.ShowDialog();
        }

        //private void comboBox18_TextUpdate(object sender, EventArgs e)
        //{
        //    this.predmetTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Predmet, comboBox18.Text);
        //}
    }
}
