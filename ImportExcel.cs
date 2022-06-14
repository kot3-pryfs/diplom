using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Monitoring_performance
{
    public partial class ImportExcel : Form
    {
        public ImportExcel()
        {
            InitializeComponent();
        }

        private void ImportExcel_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.God". При необходимости она может быть перемещена или удалена.
            this.godTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.God);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Vid_Monitoringa". При необходимости она может быть перемещена или удалена.
            this.vid_MonitoringaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Vid_Monitoringa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Predmet". При необходимости она может быть перемещена или удалена.
            this.predmetTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Predmet);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Prepod". При необходимости она может быть перемещена или удалена.
            this.prepodTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Prepod);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet1.Monitoring_View". При необходимости она может быть перемещена или удалена.
            this.monitoring_ViewTableAdapter.Fill(this.мониторинг_успеваемостиDataSet1.Monitoring_View);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Gryppa". При необходимости она может быть перемещена или удалена.
            this.gryppaTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Gryppa);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Monitoring_ImportExcel". При необходимости она может быть перемещена или удалена.
            this.monitoring_ImportExcelTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Monitoring_ImportExcel);

        }
    }
}
