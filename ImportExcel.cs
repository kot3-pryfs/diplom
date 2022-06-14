using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Monitoring_performance
{
    public partial class ImportExcel : Form
    {
        public string FIO;
        public string who;
        public bool yes;
        protected int cod;
        public ImportExcel()
        {
            InitializeComponent();
        }

        private void ImportExcel_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring". При необходимости она может быть перемещена или удалена.
            this.monitoringTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Monitoring);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Predmet_Prepod". При необходимости она может быть перемещена или удалена.
            this.predmet_PrepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Predmet_Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Student". При необходимости она может быть перемещена или удалена.
            this.studentTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Student);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.God". При необходимости она может быть перемещена или удалена.
            this.godTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.God);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Vid_Monitoringa". При необходимости она может быть перемещена или удалена.
            this.vid_MonitoringaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Vid_Monitoringa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Prepod". При необходимости она может быть перемещена или удалена.
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Gryppa". При необходимости она может быть перемещена или удалена.
            this.gryppaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Gryppa);
            comboBox1.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "Экспортировать")
            {
                int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataBySG(this.dataGridView1.Rows[1].Cells[2].Value.ToString()).Rows[0]["IdGryppa"]);
                int RowCount = Convert.ToInt32(this.studentTableAdapter.ScalarQuery(idGryppa));
                int CountPredmet = dataGridView1.RowCount / RowCount;
                MessageBox.Show(CountPredmet.ToString());

                if (comboBox1.Text != "" && comboBox5.Text != "" && comboBox6.Text != "")
                {
                    if (this.dataGridView1.Rows.Count == 0)
                    {
                        MessageBox.Show("Нет данных для выгрузки в Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (MessageBox.Show("Выгрузить найденные строки в Excel?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        return;

                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Workbook xlWB;
                    Worksheet xlSht;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWB = xlApp.Workbooks.Add();
                    xlSht = xlWB.Worksheets[1]; //первый по порядку лист в книге Excel

                    int RowsCount = this.dataGridView1.RowCount / CountPredmet;
                    int ColumnsCount = this.dataGridView1.ColumnCount + CountPredmet - 11;
                    object[,] arrData = new object[RowsCount, ColumnsCount];
                    int j1 = 0;
                    int x1 = 0;

                    for (int j = 0; j < RowsCount; j++)
                    {
                        x1 = 0;
                        for (int x = 0; x < ColumnsCount; x++)
                        {
                            if (j != this.dataGridView1.NewRowIndex)
                            {
                                if (x > 1)
                                {
                                    arrData[j1, x1] = this.dataGridView1.Rows[j + (RowCount * (x - 1))].Cells[1].Value.ToString();
                                }
                                else
                                {
                                    arrData[j1, x1] = this.dataGridView1.Rows[j].Cells[x].Value.ToString();
                                }


                            }
                            x1++;
                        }
                        j1++;
                    }

                    //выгрузка данных на лист Excel
                    xlSht.Range["B4"].Resize[arrData.GetUpperBound(0) + 1, arrData.GetUpperBound(1) + 1].Value = arrData;

                    //переносим названия столбцов в Excel файл
                    for (int j = 0; j < this.dataGridView1.Columns.Count - 10; j++)
                    {
                        xlSht.Cells[3, j + 2] = this.dataGridView1.Columns[j].HeaderCell.Value.ToString();
                    }


                    xlSht.Cells[1, 3] = this.dataGridView1.Rows[1].Cells[3].Value.ToString();
                    xlSht.Cells[2, 3] = this.dataGridView1.Rows[1].Cells[4].Value.ToString();

                    //украшательство таблицы
                    xlSht.Cells[3, 1].CurrentRegion.Borders.LineStyle = XlLineStyle.xlContinuous; //границы
                    xlSht.Rows[2].Font.Bold = true;
                    xlSht.Range["A:H"].EntireColumn.AutoFit();
                    //отображаем Excel
                    xlApp.Visible = true;
                }
                else
                {
                    MessageBox.Show("Не заполнены обязательные поля поиска", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            else
            {
                int rowcount = dataGridView1.CurrentCell.RowIndex;
                int idMonitoring = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[7].Value);
                int Ocenka = Convert.ToInt32(dataGridView1.Rows[rowcount].Cells[1].Value);
                this.monitoringTableAdapter.UpdateQuery(Ocenka, idMonitoring);
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void comboBox5_TextUpdate(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void comboBox6_TextUpdate(object sender, EventArgs e)
        {
            switch (who)
            {
                case "Заведующий отделением":
                    {
                        // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                        this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                        break;
                    }
                case "Куратор":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                            this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        break;
                    }
                case "Учащийся":
                    {
                        try
                        {
                            string[] buf = FIO.Split(' ');
                            int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                            this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;
                    }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button2.Text == "Экспортировать")
            {
                comboBox1.SelectedIndex = -1;
                comboBox5.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
                comboBox1.Text = "";
                comboBox5.Text = "";
                comboBox6.Text = "";
                switch (who)
                {
                    case "Заведующий отделением":
                        {
                            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                            this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                            break;
                        }
                    case "Куратор":
                        {
                            try
                            {
                                string[] buf = FIO.Split(' ');
                                int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                                this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                            }
                            catch
                            {
                                MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                            break;
                        }
                    case "Учащийся":
                        {
                            try
                            {
                                string[] buf = FIO.Split(' ');
                                int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                                this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                            }
                            catch
                            {
                                MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            break;
                        }
                }
            }
            else
            {
                button2.Text = "Экспортировать";
                cod = 0;
                comboBox1.SelectedIndex = -1;
                comboBox5.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
                comboBox1.Text = "";
                comboBox5.Text = "";
                comboBox6.Text = "";
                switch (who)
                {
                    case "Заведующий отделением":
                        {
                            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
                            this.monitoring_ImportExcelTableAdapter.FillBySearch(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox1.Text, comboBox5.Text, comboBox6.Text);
                            break;
                        }
                    case "Куратор":
                        {
                            try
                            {
                                string[] buf = FIO.Split(' ');
                                int idGryppa = Convert.ToInt32(this.gryppaTableAdapter.GetDataByKyrator(buf[0], buf[1], buf[2]).Rows[0]["IdGryppa"]);
                                this.monitoring_ImportExcelTableAdapter.FillByKyrator(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, comboBox5.Text, comboBox6.Text, idGryppa);
                            }
                            catch
                            {
                                MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                            break;
                        }
                    case "Учащийся":
                        {
                            try
                            {
                                string[] buf = FIO.Split(' ');
                                int idStudent = Convert.ToInt32(this.studentTableAdapter.GetDataByStudent(buf[0], buf[1], buf[2]).Rows[0]["IdStudent"]);
                                this.monitoring_ImportExcelTableAdapter.FillByStudent(мониторинг_успеваемостиDataSet.Monitoring_ImportExcel, idStudent, comboBox5.Text, comboBox6.Text);
                            }
                            catch
                            {
                                MessageBox.Show("Ошибка поиска данных", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }
                            break;
                        }
                }
            }

        }

        private void splitContainer1_Panel1_Click(object sender, EventArgs e)
        {
            if (yes == true)
            {
                cod++;
                string pass = "";
                if (cod > 5)
                {
                    pass = comboBox1.Text + comboBox5.Text;
                }
                if (pass == "REDACTMY")
                {
                    string message = "Вы уверены что хотите активировать режим редактирования?";
                    string caption = "Мониторинг успеваемости";
                    MessageBoxButtons button = MessageBoxButtons.YesNo;
                    MessageBoxIcon icon = MessageBoxIcon.Warning;
                    DialogResult result = MessageBox.Show(message, caption, button, icon);
                    if (result == DialogResult.Yes)
                    {
                        label2.Visible = true;
                        comboBox2.Visible = true;
                        button2.Text = "Применить";
                    }
                }
            }
        }

    }
}
