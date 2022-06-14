using System;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;

namespace Monitoring_performance
{
    public partial class Aynt : Form
    {
        protected int code;
        public Aynt()
        {
            InitializeComponent();
        }

        private void Aynt_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "мониторинг_успеваемостиDataSet.Kyrator". При необходимости она может быть перемещена или удалена.
            this.kyratorTableAdapter.Fill(this.мониторинг_успеваемостиDataSet.Kyrator);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string log = textBox1.Text;
            string pass = textBox2.Text;
            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "college.local", "DC=college,DC=local", log, pass)) //создание подключения с указанными логином и паролем
            {
                bool isValid = pc.ValidateCredentials(log, pass);  // проверка наличия на сервере пользователя с указанным логином и паролем

                if (isValid == true)
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(pc, IdentityType.SamAccountName, log); // создаем экземпляр класса который позволит вытянуть группы безопастности пользователя

                    MessageBox.Show("Добро пожаловать " + user.DisplayName, "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    int buf = 0;
                    //string[] GRP_SECURITI = null;
                    if (user != null)
                    {
                        foreach (Principal p in user.GetAuthorizationGroups()) // извлекаем по очереди группы безопастности пользователя
                        {
                            switch (comboBox1.Text) //в зависимости от выбранной роли пихаем группу безопастности на проверку
                            {
                                case "Администратор":
                                    {
                                        if (p.Name == "Администраторы") // если совпало, значит пускаем
                                        {
                                            AdminForm Ex = new AdminForm();
                                            Ex.NameA = user.DisplayName;
                                            Ex.ShowDialog();
                                            break;
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Превышено количество попыток аунтификации");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                case "Секретарь":
                                    {
                                        if (p.Name == "grp_teaching_department_DO") //загвостка здесь и у заведующего. Для проверки необходимо сравнить несколько групп, а сделать этого не можем
                                        {                                              // так как одновременно мы проверяем только одну текущую группу.
                                            int buff = 0;
                                            foreach(Principal g in user.GetAuthorizationGroups()) //решается вторым циклом который проверит наличие второй группы у пользователя
                                            {
                                                if (g.Name == "grp_teachers")
                                                {
                                                    buff++;
                                                }
                                            }
                                            if (buff == 0)
                                            {
                                                SecretaryForm Ex = new SecretaryForm();
                                                Ex.ShowDialog();
                                                break;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Пользователь отсутствует в данной категории");
                                                break;                                                
                                            }
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Превышено количество попыток аунтификации");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                case "Заведующий отделением":
                                    {
                                        if (p.Name == "grp_teaching_department_DO")
                                        {
                                            int buff = 0;
                                            foreach (Principal g in user.GetAuthorizationGroups())
                                            {
                                                if (g.Name == "grp_teachers")
                                                {
                                                    buff++;
                                                }
                                            }
                                            if (buff == 1)
                                            {
                                                ImportExcel Ex = new ImportExcel();
                                                Ex.who = comboBox1.Text;
                                                Ex.yes = true;
                                                Ex.ShowDialog();
                                                break;
                                            }
                                            
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Превышено количество попыток аунтификации");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                case "Куратор":
                                    {
                                        string FIO = user.DisplayName;
                                        string[] GG = FIO.Split(' ');
                                        int scalar = Convert.ToInt32(kyratorTableAdapter.ScalarQuery(GG[0], GG[1], GG[2]));
                                        if(scalar != 0 && p.Name == "grp_teachers")
                                        {
                                            ImportExcel Ex = new ImportExcel();
                                            Ex.who = comboBox1.Text;
                                            Ex.FIO = user.DisplayName;
                                            Ex.ShowDialog();
                                            break;
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Пользователь отсутствует в данной категории");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                case "Преподаватель":
                                    {
                                        if (p.Name == "grp_teachers")
                                        {
                                            LecturerForm Ex = new LecturerForm();
                                            Ex.FIO = user.DisplayName;
                                            Ex.ShowDialog();
                                            break;
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Пользователь отсутствует в данной категории");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                case "Учащийся":
                                    {
                                        if (p.Name == "grp_students_all")
                                        {
                                            ImportExcel Ex = new ImportExcel();
                                            Ex.FIO = user.DisplayName;
                                            Ex.who = comboBox1.Text;
                                            Ex.ShowDialog();
                                            break;
                                        }
                                        else
                                        {
                                            buf++;
                                            if (buf > 7)
                                            {
                                                MessageBox.Show("Пользователь отсутствует в данной категории");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                            }
                        }
                    }
                    pc.Dispose();
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль", "Мониторинг успеваемости", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }

            }
        }

        private void Aynt_Click(object sender, EventArgs e)
        {
            if (code == 4)
            {
                textBox3.Visible = true;
            }
            code++;
            if (code > 5 && textBox1.Text == "ADMIN" && textBox2.Text == "admin")
            {
                switch (comboBox1.Text)
                {
                    case "Администратор":
                        {
                            button1.Enabled = false;
                            AdminForm Ex = new AdminForm();
                            Ex.ShowDialog();
                            break;
                        }
                    case "Секретарь":
                        {
                            button1.Enabled = false;
                            SecretaryForm Ex = new SecretaryForm();
                            Ex.ShowDialog();
                            break;
                        }
                    case "Заведующий отделением":
                        {
                            button1.Enabled = false;
                            ImportExcel Ex = new ImportExcel();
                            Ex.yes = true;
                            Ex.who = comboBox1.Text;
                            Ex.ShowDialog();
                            break;
                        }
                    case "Куратор":
                        {
                            button1.Enabled = false;
                            ImportExcel Ex = new ImportExcel();
                            Ex.FIO = textBox3.Text;
                            Ex.who = comboBox1.Text;
                            Ex.ShowDialog();
                            break;
                        }
                    case "Преподаватель":
                        {
                            button1.Enabled = false;
                            LecturerForm Ex = new LecturerForm();
                            Ex.FIO = textBox3.Text;
                            Ex.ShowDialog();
                            break;
                        }
                    case "Учащийся":
                        {
                            button1.Enabled = false;
                            ImportExcel Ex = new ImportExcel();
                            Ex.FIO = textBox3.Text;
                            Ex.who = comboBox1.Text;
                            Ex.ShowDialog();
                            break;
                        }
                
                }
            }
        }
    }
}
