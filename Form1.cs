using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TEO
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Concurent.RowCount = 9;
            Concurent.Rows[0].Cells[0].Value = "1.Удобство работы(пользовательский интерфейс)";
            Concurent.Rows[1].Cells[0].Value = "2.Новизна (соответствие современным требованиям)";
            Concurent.Rows[2].Cells[0].Value = "3.Соответствие профилю деятельности заказчика";
            Concurent.Rows[3].Cells[0].Value = "4.Ресурсная эффективность";
            Concurent.Rows[4].Cells[0].Value = "5. Надежность (защита данных)";
            Concurent.Rows[5].Cells[0].Value = "6.Скорость доступа к данным";
            Concurent.Rows[6].Cells[0].Value = "7.Гибкость настройки";
            Concurent.Rows[7].Cells[0].Value = "8.Обучаемость персонала";
            Concurent.Rows[8].Cells[0].Value = "9.Соотношение стоимость/возможности";
            for (int i = 0; i <= 7; i++)
            {
                Concurent.Rows[i].Cells[1].Value = 0.1;
            }
            Concurent.Rows[8].Cells[1].Value = 0.2;
            //////////////////////////////////////////////////////////////////////////////
            Calendar.RowCount = 36;
            Calendar.Rows[0].Cells[0].Value = "1.1 Постановка задачи";
            Calendar.Rows[0].Cells[1].Value = "Руководитель";
            Calendar.Rows[1].Cells[1].Value = "Программист";
            Calendar.Rows[2].Cells[0].Value = "1.2 Сбор исходных данных";
            Calendar.Rows[2].Cells[1].Value = "Руководитель";
            Calendar.Rows[3].Cells[1].Value = "Программист";
            Calendar.Rows[4].Cells[0].Value = "1.3 Анализ существующих методов решения задачи и программных средств";
            Calendar.Rows[4].Cells[1].Value = "Руководитель";
            Calendar.Rows[5].Cells[1].Value = "Программист";
            Calendar.Rows[6].Cells[0].Value = "1.4 Обоснование принципиальной необходимости разработки";
            Calendar.Rows[6].Cells[1].Value = "Руководитель";
            Calendar.Rows[7].Cells[1].Value = "Программист";
            Calendar.Rows[8].Cells[0].Value = "1.5 Определение и анализ требований к программе";
            Calendar.Rows[8].Cells[1].Value = "Руководитель";
            Calendar.Rows[9].Cells[1].Value = "Программист";
            Calendar.Rows[10].Cells[0].Value = "1.6 Определение структуры входных и выходных данных";
            Calendar.Rows[10].Cells[1].Value = "Руководитель";
            Calendar.Rows[11].Cells[1].Value = "Программист";
            Calendar.Rows[12].Cells[0].Value = "1.7 Выбор технических средств и программных средств реализации";
            Calendar.Rows[12].Cells[1].Value = "Руководитель";
            Calendar.Rows[13].Cells[1].Value = "Программист";
            Calendar.Rows[14].Cells[0].Value = "1.8 Согласование и утверждение технического задания";
            Calendar.Rows[14].Cells[1].Value = "Руководитель";
            Calendar.Rows[15].Cells[1].Value = "Программист";
            Calendar.Rows[16].Cells[0].Value = "2.1 Проектирование программной архитектуры";
            Calendar.Rows[16].Cells[1].Value = "Руководитель";
            Calendar.Rows[17].Cells[1].Value = "Программист";
            Calendar.Rows[18].Cells[0].Value = "2.2 Техническое проектирование компонентов программы";
            Calendar.Rows[18].Cells[1].Value = "Руководитель";
            Calendar.Rows[19].Cells[1].Value = "Программист";
            Calendar.Rows[20].Cells[0].Value = "3.1 Программирование модулей в выбранной среде программирования";
            Calendar.Rows[20].Cells[1].Value = "Руководитель";
            Calendar.Rows[21].Cells[1].Value = "Программист";
            Calendar.Rows[22].Cells[0].Value = "3.2 Тестирование программных модулей";
            Calendar.Rows[22].Cells[1].Value = "Руководитель";
            Calendar.Rows[23].Cells[1].Value = "Программист";
            Calendar.Rows[24].Cells[0].Value = "3.3  Сборка и испытание программы";
            Calendar.Rows[24].Cells[1].Value = "Руководитель";
            Calendar.Rows[25].Cells[1].Value = "Программист";
            Calendar.Rows[26].Cells[0].Value = "3.4 Анализ результатов испытаний";
            Calendar.Rows[26].Cells[1].Value = "Руководитель";
            Calendar.Rows[27].Cells[1].Value = "Программист";
            Calendar.Rows[28].Cells[0].Value = "4.1  Проведение расчетов показателей безопасности жизнедеятельности";
            Calendar.Rows[28].Cells[1].Value = "Руководитель";
            Calendar.Rows[29].Cells[1].Value = "Программист";
            Calendar.Rows[30].Cells[0].Value = "4.2  Проведение экономических расчетов";
            Calendar.Rows[30].Cells[1].Value = "Руководитель";
            Calendar.Rows[31].Cells[1].Value = "Программист";
            Calendar.Rows[32].Cells[0].Value = "4.3 Оформление пояснительной записки";
            Calendar.Rows[32].Cells[1].Value = "Руководитель";
            Calendar.Rows[33].Cells[1].Value = "Программист";
            Calendar.Rows[34].Cells[0].Value = "Количество рабочих дней";
            Calendar.Rows[35].Cells[0].Value = "Количество рабочих дней";
            Calendar.Rows[34].Cells[1].Value = "Руководитель";
            Calendar.Rows[35].Cells[1].Value = "Программист";
            for (int i = 0; i <= 33; i++)
            {
                Calendar.Rows[i].Cells[2].Value = 1;
            }
            //////////////////////////////////////////////////////////////////////////////////////////
            Osnzp.RowCount = 2;
            Osnzp.Rows[0].Cells[0].Value = "Руководитель";
            Osnzp.Rows[1].Cells[0].Value = "Программист";
            Osnzp.Rows[0].Cells[1].Value = 50000;
            Osnzp.Rows[1].Cells[1].Value = 42500;
            Razrab.RowCount = 6;
            Razrab.ColumnCount = 2;
            Razrab.Rows[0].Cells[0].Value = "Основная заработная плата ";
            Razrab.Rows[1].Cells[0].Value = "Дополнительная зарплата ";
            Razrab.Rows[2].Cells[0].Value = "Отчисления на социальные нужды  ";
            Razrab.Rows[3].Cells[0].Value = "Затраты на материалы ";
            Razrab.Rows[4].Cells[0].Value = "Затраты на машинное время  ";
            Razrab.Rows[5].Cells[0].Value = "Накладные расходы организации  ";
            //////////////////////////////////////////////////////////////////////////////////////////
            zAnalog.RowCount = 4;
            zAnalog.Rows[0].Cells[0].Value = "Приобретение программного продукта ";
            zAnalog.Rows[1].Cells[0].Value = "Оплата услуг на установку и сопровождение продукта  ";
            zAnalog.Rows[2].Cells[0].Value = "Основное и вспомогательное оборудование  ";
            zAnalog.Rows[3].Cells[0].Value = "Подготовка пользователя  ";
            //////////////////////////////////////////////////////////////////////////////////////////
            DateProject.RowCount = 3;
            DateProject.Rows[0].Cells[0].Value = "Сотрудник отдела-эксплуатанта";
            DateProject.Rows[1].Cells[0].Value = "Программист";
            DateProject.Rows[2].Cells[0].Value = "";
            DateProject.Rows[0].Cells[1].Value = 35000;
            DateProject.Rows[1].Cells[1].Value = 42500;
            DateProject.Rows[0].Cells[3].Value = 10;
            DateProject.Rows[1].Cells[3].Value = 10;
            DateAnalog.RowCount = 3;
            DateAnalog.Rows[0].Cells[0].Value = "Сотрудник отдела-эксплуатанта";
            DateAnalog.Rows[1].Cells[0].Value = "Программист";
            DateAnalog.Rows[2].Cells[0].Value = "";
            DateAnalog.Rows[0].Cells[1].Value = 35000;
            DateAnalog.Rows[1].Cells[1].Value = 42500;
            DateAnalog.Rows[0].Cells[3].Value = 17;
            DateAnalog.Rows[1].Cells[3].Value = 17;
            //////////////////////////////////////////////////////////////////////////////////////////
            EcEff.RowCount = 3;
            EcEff.Rows[0].Cells[0].Value = "Себестоимость (текущие эксплуатационные затраты), руб. ";
            EcEff.Rows[1].Cells[0].Value = "Суммарные затраты, связанные с внедрением проекта, руб. ";
            EcEff.Rows[2].Cells[0].Value = "Приведенные затраты на единицу работ, руб. ";
            ResEff.RowCount = 5;
            ResEff.Rows[0].Cells[0].Value = "Затраты на разработку и внедрение проекта, руб. ";
            ResEff.Rows[1].Cells[0].Value = "Общие эксплуатационные затраты, руб. ";
            ResEff.Rows[2].Cells[0].Value = "Экономический эффект, руб. ";
            ResEff.Rows[3].Cells[0].Value = "Коэффициент экономической эффективности ";
            ResEff.Rows[4].Cells[0].Value = "Срок окупаемости, лет ";
            //////////////////////////////////////////////////////////////////////////////////////////
            YearZ.RowCount = 7;
            YearZ.Rows[0].Cells[0].Value = "Основная и дополнительная зарплата с \n отчислениями во внебюджетные фонды ";
            YearZ.Rows[1].Cells[0].Value = "Амортизационные отчисления ";
            YearZ.Rows[2].Cells[0].Value = "Затраты на электроэнергию ";
            YearZ.Rows[3].Cells[0].Value = "Затраты на текущий ремонт ";
            YearZ.Rows[4].Cells[0].Value = "Затраты на материалы ";
            YearZ.Rows[5].Cells[0].Value = "Накладные расходы ";
            YearZ.Rows[6].Cells[0].Value = "Итого ";
            
            zAnalog.Rows[0].Cells[0].Value = "Затраты на приобретение программного продукта";
            zAnalog.Rows[0].Cells[1].Value = 37300;
            zAnalog.Rows[1].Cells[0].Value = "Затраты по оплате услуг на установку и сопровождение продукта";
            zAnalog.Rows[1].Cells[1].Value = 12000;
            zAnalog.Rows[2].Cells[0].Value = "Затраты на основное и вспомогательное оборудование";
            zAnalog.Rows[2].Cells[1].Value = 22500;
            zAnalog.Rows[3].Cells[0].Value = "Затраты на подготовку пользователя ";
            zAnalog.Rows[3].Cells[1].Value = 9000;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            double kf = 0;
            foreach (DataGridViewRow row in Concurent.Rows)
            {
                double kf1 = 0;
                double.TryParse((row.Cells[1].Value ?? "0").ToString().Replace(".", ","), out kf1);
                //double kf1 = Convert.ToDouble(row.Cells[1].Value);
                kf += kf1;
                
            }
            
            if (kf == 1)
            {
                for (int i = 0; i < Concurent.Rows.Count; i++)
                {
                    // MessageBox.Show(Convert.ToString(Concurent[2, i].Value) + " and " + Convert.ToString(Concurent[4, i].Value));
                    double a;
                    double.TryParse((Concurent[1, i].Value ?? "0").ToString().Replace(".", ","), out a);
                    if (a != 0)
                    {

                        if ((Convert.ToInt32(Concurent[2, i].Value) != 0 && Convert.ToInt32(Concurent[4, i].Value) != 0))
                        {

                            double kf2;
                            double.TryParse((Concurent[1, i].Value ?? "0").ToString().Replace(".", ","), out kf2);

                            Concurent[3, i].Value = kf2 * Convert.ToDouble(Concurent[2, i].Value);
                            Concurent[5, i].Value = kf2 * Convert.ToDouble(Concurent[4, i].Value);

                        }
                        else if ((Convert.ToInt32(Concurent[2, i].Value) == 0 || Convert.ToInt32(Concurent[4, i].Value) == 0))
                            MessageBox.Show("Не введены экспертные оценки проекта и/или его конкурента в " + (i + 1) + " строке");
                    }
                }
                double jety1 = 0, jety2 = 0;
                foreach (DataGridViewRow row in Concurent.Rows)
                {
                    double s;
                    double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out s);
                    jety1 += s;

                }

                textBox1.Text = Convert.ToString(jety1);
                foreach (DataGridViewRow row in Concurent.Rows)
                {
                    double s;
                    double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out s);
                    jety2 += s;

                }

                textBox2.Text = Convert.ToString(jety2);
                double A = Math.Round(jety1 / jety2, 2);
                textBox3.Text = Convert.ToString(A);
                if (A > 1) MessageBox.Show("Разработка проекта оправдана");
                else MessageBox.Show("Разработка проекта не оправдана");
            }
            else MessageBox.Show("Сумма коэффициентов не равна 1");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime date1 = new DateTime();
            DateTime dateB, dateM;
            bool allow = false;
            for(int i = 0; i <= 33; i++)
            {
                if(Calendar.Rows[i].Cells[2].Value == null)
                {
                    allow = false;
                    MessageBox.Show("Не все дни заполнены");
                    break;
                }
                else { allow = true; }
            }
            if(allow == true)
            {
                Calendar.Rows[0].Cells[3].Value = dateTimePicker1.Value.ToString("dd.MM.yyyy");
                Calendar.Rows[1].Cells[3].Value = Calendar.Rows[0].Cells[3].Value;

                date1 = Convert.ToDateTime(Calendar.Rows[0].Cells[3].Value); 
                date1 = date1.AddDays(Convert.ToInt32(Calendar.Rows[0].Cells[2].Value) - 1); 
                Calendar.Rows[0].Cells[4].Value = date1.ToShortDateString();

                date1 = Convert.ToDateTime(Calendar.Rows[1].Cells[3].Value); 
                date1 = date1.AddDays(Convert.ToInt32(Calendar.Rows[1].Cells[2].Value) - 1); 
                Calendar.Rows[1].Cells[4].Value = date1.ToShortDateString(); 
            }
            for (int i = 2; i <= 33; i = i + 2)
            {
                int d = Convert.ToInt32(Calendar.Rows[i].Cells[2].Value);

                if (d == 0)
                {
                    Calendar.Rows[i].Cells[3].Value = null; Calendar.Rows[i].Cells[4].Value = null;
                    dateM = Convert.ToDateTime(Calendar.Rows[i - 2].Cells[4].Value);
                    dateB = Convert.ToDateTime(Calendar.Rows[i - 2].Cells[4].Value);
                    Calendar.Rows[i].Cells[3].Value = "-";
                    Calendar.Rows[i].Cells[4].Value = "-";

                    if (dateB > dateM)
                        date1 = Convert.ToDateTime(Calendar.Rows[i - 2].Cells[4].Value);
                    else date1 = Convert.ToDateTime(Calendar.Rows[i - 2].Cells[4].Value);
                }
                else
                {
                    dateM = Convert.ToDateTime(Calendar.Rows[i - 1].Cells[4].Value);
                    dateB = Convert.ToDateTime(Calendar.Rows[i - 1].Cells[4].Value);
                    if (dateB > dateM)
                        date1 = Convert.ToDateTime(Calendar.Rows[i - 1].Cells[4].Value);
                    else date1 = Convert.ToDateTime(Calendar.Rows[i - 1].Cells[4].Value);
                    date1 = date1.AddDays(1);
                    Calendar.Rows[i].Cells[3].Value = date1.ToShortDateString();
                    date1 = date1.AddDays(Convert.ToInt32(Calendar.Rows[i].Cells[2].Value) - 1);
                    Calendar.Rows[i].Cells[4].Value = date1.ToShortDateString();

                    Calendar.Rows[i + 1].Cells[3].Value = Calendar.Rows[i].Cells[3].Value;

                    date1 = Convert.ToDateTime(Calendar.Rows[i + 1].Cells[3].Value);
                    date1 = date1.AddDays(Convert.ToInt32(Calendar.Rows[i + 1].Cells[2].Value) - 1);
                    Calendar.Rows[i + 1].Cells[4].Value = date1.ToShortDateString();
                }
            }
                int daycount1 = 0;
                int daycount2 = 0;
                for (int i = 0; i <= 33; i = i + 2)
                {

                    daycount1 += Convert.ToInt32(Calendar.Rows[i].Cells[2].Value);
                    daycount2 += Convert.ToInt32(Calendar.Rows[i + 1].Cells[2].Value);
                }

                Calendar.Rows[34].Cells[2].Value = daycount1;
                Calendar.Rows[35].Cells[2].Value = daycount2;
            }

        private void Button3_Click(object sender, EventArgs e)
        {
            if ((RM.Text == String.Empty) || (KO.Text == String.Empty) || (KS.Text == String.Empty) || (KN.Text == String.Empty) || (RK.Text == String.Empty) || (Pr.Text == String.Empty) || (CN.Text == String.Empty) || (_1hp.Text == String.Empty) || (pcT.Text == String.Empty) || (kmp.Text == String.Empty) || (_1ht.Text == String.Empty) || (dvn.Text == String.Empty) || (hpd.Text == String.Empty) || (dpy.Text == String.Empty)) { MessageBox.Show("заполните поля"); }
            else
            {
                Osnzp.Rows[0].Cells[3].Value = Convert.ToInt32(Calendar.Rows[34].Cells[2].Value);
                Osnzp.Rows[1].Cells[3].Value = Convert.ToInt32(Calendar.Rows[35].Cells[2].Value);
                double del;
                double j2 = double.Parse(RM.Text.Replace(".", ","));
                del = Convert.ToDouble(Osnzp.Rows[0].Cells[1].Value) / j2;
                del = Math.Round(del, 2);
                Osnzp.Rows[0].Cells[2].Value = Convert.ToDouble(del);

                del = Convert.ToDouble(Osnzp.Rows[1].Cells[1].Value) / j2;
                del = Math.Round(del, 2);
                Osnzp.Rows[1].Cells[2].Value = Convert.ToDouble(del);

                Osnzp.Rows[0].Cells[4].Value = Convert.ToDouble(Osnzp.Rows[0].Cells[2].Value) * Convert.ToDouble(Osnzp.Rows[0].Cells[3].Value);
                Osnzp.Rows[1].Cells[4].Value = Convert.ToDouble(Osnzp.Rows[1].Cells[2].Value) * Convert.ToDouble(Osnzp.Rows[1].Cells[3].Value);
                double j1 = Convert.ToSingle(Osnzp.Rows[0].Cells[4].Value) + Convert.ToSingle(Osnzp.Rows[1].Cells[4].Value);
                j1 = Math.Round(j1, 2);
                it1.Text = Convert.ToString(j1);

                double j22 = 0;
                for (int i = 0; i < Mat.Rows.Count - 1; i++)
                {
                    Mat[4, i].Value = Convert.ToDouble(Mat[2, i].Value) * Convert.ToDouble(Mat[3, i].Value);

                    j22 += Convert.ToDouble(Mat.Rows[i].Cells[4].Value);
                }
                j22 = Math.Round(j22, 2);
                it3.Text = Convert.ToString(j22);

                Razrab.Rows[0].Cells[1].Value = Convert.ToString(j1);

                double z, m, n;
                z = Convert.ToDouble(Razrab.Rows[0].Cells[1].Value) * (Convert.ToDouble(KO.Text) + Convert.ToDouble(RK.Text));
                z = Math.Round(z, 2);
                Razrab.Rows[1].Cells[1].Value = Convert.ToDouble(z);

                double s = double.Parse(KS.Text.Replace(".", ","));
                Razrab.Rows[2].Cells[1].Value = (Convert.ToDouble(Razrab.Rows[0].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[1].Cells[1].Value)) * s;

                Razrab.Rows[3].Cells[1].Value = Convert.ToString(j22);

                m = Convert.ToDouble(_1hp.Text) * Convert.ToDouble(pcT.Text) * Convert.ToDouble(kmp.Text) * 115;
                m = Math.Round(m, 2);
                Razrab.Rows[4].Cells[1].Value = Convert.ToDouble(m);

                n = Convert.ToDouble(Razrab.Rows[0].Cells[1].Value) * Convert.ToDouble(KN.Text);
                n = Math.Round(n, 2);
                Razrab.Rows[5].Cells[1].Value = Convert.ToDouble(n);

                double ss = Convert.ToDouble(Razrab.Rows[0].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[1].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[2].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[3].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[4].Cells[1].Value) + Convert.ToDouble(Razrab.Rows[5].Cells[1].Value);
                ss = Math.Round(ss, 2);
                it2.Text = Convert.ToString(ss);

                double kp = (Convert.ToDouble(Pr.Text) * Convert.ToDouble(CN.Text) * Convert.ToDouble(_1ht.Text) * Convert.ToDouble(dvn.Text)) / (Convert.ToDouble(dpy.Text) * Convert.ToDouble(hpd.Text));
                kp = Math.Round(kp, 2);
                cRel.Text = Convert.ToString(kp);

                double kap = Convert.ToDouble(ss) + Convert.ToDouble(kp);
                cIn.Text = Convert.ToString(kap);

                double an = 0;
                for (int i = 0; i < zAnalog.Rows.Count - 1; i++)
                {
                    an += Convert.ToDouble(zAnalog.Rows[i].Cells[1].Value);
                }
                an = Math.Round(an, 2);
                pAn.Text = Convert.ToString(an);

            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            double j = 0, v, sp = 0, sa = 0;
            if ((dpm1.Text == String.Empty) || (bPr.Text == String.Empty) || (cpp.Text == String.Empty) || (Amor.Text == String.Empty) || (NL.Text == String.Empty) || (bAn.Text == String.Empty) || (kAn.Text == String.Empty) || (KO.Text == String.Empty) || (RO1.Text == String.Empty) || (normR.Text == String.Empty) || (_1ht.Text == String.Empty) || (pKVTh.Text == String.Empty) || (KPU.Text == String.Empty) || (taxE.Text == String.Empty)) { MessageBox.Show("заполните поля"); }
            else
            {
                for (int i = 0; i < DateProject.Rows.Count - 1; i++)
                {
                    sp += Convert.ToDouble(DateProject[3, i].Value);
                    v = Convert.ToDouble(DateProject[1, i].Value) / Convert.ToDouble(dpm1.Text);
                    v = Math.Round(v, 2);
                    DateProject[2, i].Value = Convert.ToDouble(v);
                    v = Convert.ToDouble(DateProject[2, i].Value) * Convert.ToDouble(DateProject[3, i].Value) * (1 + Convert.ToDouble(RO1.Text) + Convert.ToDouble(KO.Text)) * (1 + Convert.ToDouble(KS.Text));
                    v = Math.Round(v, 2);
                    DateProject[4, i].Value = Convert.ToDouble(v);
                    j += Convert.ToDouble(DateProject.Rows[i].Cells[4].Value);
                }
                j = Math.Round(j, 2);
                it_1.Text = Convert.ToString(j);

                j = 0;
                for (int i = 0; i < DateAnalog.Rows.Count - 1; i++)
                {
                    sa += Convert.ToDouble(DateAnalog[3, i].Value);
                    v = Convert.ToDouble(DateAnalog[1, i].Value) / Convert.ToDouble(dpm1.Text);
                    v = Math.Round(v, 2);
                    DateAnalog[2, i].Value = Convert.ToDouble(v);
                    v = Convert.ToDouble(DateAnalog[2, i].Value) * Convert.ToDouble(DateAnalog[3, i].Value) * (1 + Convert.ToDouble(RO1.Text) + Convert.ToDouble(KO.Text)) * (1 + Convert.ToDouble(KS.Text));
                    v = Math.Round(v, 2);
                    DateAnalog[4, i].Value = Convert.ToDouble(v);
                    j += Convert.ToDouble(DateAnalog.Rows[i].Cells[4].Value);
                }
                j = Math.Round(j, 2);
                it_3.Text = Convert.ToString(j);

                YearZ.Rows[0].Cells[1].Value = Convert.ToDouble(it_1.Text);
                YearZ.Rows[0].Cells[2].Value = Convert.ToDouble(it_3.Text);

                v = (Convert.ToDouble(bPr.Text) * Convert.ToDouble(Amor.Text) * Convert.ToDouble(cpp.Text) * (sp * Convert.ToDouble(NL.Text))) / (247 * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[1].Cells[1].Value = Convert.ToDouble(v);

                v = (Convert.ToDouble(bAn.Text) * Convert.ToDouble(Amor.Text) * Convert.ToDouble(kAn.Text) * (sa * Convert.ToDouble(NL.Text))) / (247 * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[1].Cells[2].Value = Convert.ToDouble(v);

                v = Convert.ToDouble(pKVTh.Text) * Convert.ToDouble(KPU.Text) * Convert.ToDouble(taxE.Text) * (sp * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[2].Cells[1].Value = Convert.ToDouble(v);

                v = Convert.ToDouble(pKVTh.Text) * Convert.ToDouble(KPU.Text) * Convert.ToDouble(taxE.Text) * (sa * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[2].Cells[2].Value = Convert.ToDouble(v);

                v = (Convert.ToDouble(normR.Text) * (Convert.ToDouble(bPr.Text) * sp * Convert.ToDouble(NL.Text))) / (Convert.ToDouble(dpy.Text) * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[3].Cells[1].Value = Convert.ToDouble(v);

                v = (Convert.ToDouble(normR.Text) * Convert.ToDouble(bAn.Text) * sa * Convert.ToDouble(NL.Text)) / (Convert.ToDouble(dpy.Text) * Convert.ToDouble(NL.Text));
                v = Math.Round(v, 2);
                YearZ.Rows[3].Cells[2].Value = Convert.ToDouble(v);

                v = Convert.ToDouble(bPr.Text) * Convert.ToDouble(MatE.Text);
                v = Math.Round(v, 2);
                YearZ.Rows[4].Cells[1].Value = Convert.ToDouble(v);

                v = Convert.ToDouble(bAn.Text) * Convert.ToDouble(MatE.Text);
                v = Math.Round(v, 2);
                YearZ.Rows[4].Cells[2].Value = Convert.ToDouble(v);

                v = (Convert.ToDouble(YearZ.Rows[4].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[3].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[2].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[1].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[0].Cells[1].Value)) * Convert.ToDouble(NakE.Text);
                v = Math.Round(v, 2);
                YearZ.Rows[5].Cells[1].Value = Convert.ToDouble(v);

                v = (Convert.ToDouble(YearZ.Rows[4].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[3].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[2].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[1].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[0].Cells[2].Value)) * Convert.ToDouble(NakE.Text);
                v = Math.Round(v, 2);
                YearZ.Rows[5].Cells[2].Value = Convert.ToDouble(v);

                v = Convert.ToDouble(YearZ.Rows[5].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[4].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[3].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[2].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[1].Cells[1].Value) + Convert.ToDouble(YearZ.Rows[0].Cells[1].Value);
                v = Math.Round(v, 2);
                YearZ.Rows[6].Cells[1].Value = v;

                v = Convert.ToDouble(YearZ.Rows[5].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[4].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[3].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[2].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[1].Cells[2].Value) + Convert.ToDouble(YearZ.Rows[0].Cells[2].Value);
                v = Math.Round(v, 2);
                YearZ.Rows[6].Cells[2].Value = v;

            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            
                string s1 = Convert.ToString(YearZ.Rows[6].Cells[2].Value);
                string s2 = Convert.ToString(YearZ.Rows[6].Cells[1].Value);
                string s3 = Convert.ToString(pAn.Text);
                string s4 = Convert.ToString(cIn.Text);
                string s5 = Convert.ToString(normKof.Text);
                string s6 = Convert.ToString(textBox3.Text);
                if ((String.IsNullOrEmpty(s1) == true) || (String.IsNullOrEmpty(s2) == true) || (String.IsNullOrEmpty(s3) == true) || (String.IsNullOrEmpty(s4) == true) || (String.IsNullOrEmpty(s5) == true) || (String.IsNullOrEmpty(s6) == true))
                {
                    MessageBox.Show("Нет какого-то из решенных параметров");
                    return;
                }

                EcEff.Rows[0].Cells[1].Value = Convert.ToDouble(YearZ.Rows[6].Cells[2].Value);
                EcEff.Rows[0].Cells[2].Value = Convert.ToDouble(YearZ.Rows[6].Cells[1].Value);

                EcEff.Rows[1].Cells[1].Value = Convert.ToDouble(pAn.Text);
                EcEff.Rows[1].Cells[2].Value = Convert.ToDouble(cIn.Text);

                EcEff.Rows[2].Cells[1].Value = Convert.ToDouble(EcEff.Rows[0].Cells[1].Value) + Convert.ToDouble(EcEff.Rows[1].Cells[1].Value) * Convert.ToDouble(normKof.Text);
                EcEff.Rows[2].Cells[2].Value = Convert.ToDouble(EcEff.Rows[0].Cells[2].Value) + Convert.ToDouble(EcEff.Rows[1].Cells[2].Value) * Convert.ToDouble(normKof.Text);
                double l = Convert.ToDouble(textBox3.Text);
                if (l == 0)
                    MessageBox.Show("zero");
                double j = Convert.ToDouble(EcEff.Rows[2].Cells[1].Value) * l - Convert.ToDouble(EcEff.Rows[2].Cells[2].Value);
                j = Math.Round(j, 2);
                econEff.Text = Convert.ToString(j);

                ResEff.Rows[0].Cells[1].Value = Convert.ToDouble(EcEff.Rows[1].Cells[2].Value);
                ResEff.Rows[1].Cells[1].Value = Convert.ToDouble(EcEff.Rows[0].Cells[2].Value);
                ResEff.Rows[2].Cells[1].Value = Convert.ToDouble(econEff.Text);
                j = Convert.ToDouble(ResEff.Rows[0].Cells[1].Value) / Convert.ToDouble(ResEff.Rows[2].Cells[1].Value);
                j = Math.Round(j, 2);
                ResEff.Rows[4].Cells[1].Value = Convert.ToDouble(j);

                j = 1 / Convert.ToDouble(ResEff.Rows[4].Cells[1].Value);
                j = Math.Round(j, 2);
                ResEff.Rows[3].Cells[1].Value = Convert.ToDouble(j);

                if (Convert.ToDouble(ResEff.Rows[3].Cells[1].Value) > Convert.ToDouble(normKof.Text)) MessageBox.Show("Проект эффективен");
                else MessageBox.Show("Проект не эффективен, следует купить его аналог");
            
        }

        private void ОбАвтореToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Работа по дисциплине Программная Инженерия выполнена Лебедевым Иваном, студентом группы 448-2." + " ©Ju-Van, 2020", "О программе");
        }

        private void СправкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
           System.Diagnostics.Process.Start(".\\Справка.docx");
        }

        private void pKVTh_TextChanged(object sender, EventArgs e)
        {

        }

        private void лицензияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(".\\Лицензия.docx");
        }
    }
    }

