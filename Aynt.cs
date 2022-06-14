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
    public partial class Aynt : Form
    {
        public Aynt()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AdminForm Ex = new AdminForm();
            Ex.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SecretaryForm Ex = new SecretaryForm();
            Ex.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LecturerForm Ex = new LecturerForm();
            Ex.ShowDialog();
        }
    }
}
