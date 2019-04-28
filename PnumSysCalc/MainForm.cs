using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace PnumMixCalc
{
    public partial class MainForm : Form
    {
        //Инициализация переменных
        const int countOfArr = 10000000;
        double[] P = new double[4];
        List<double> PCheck = new List<double>(); //Динамический массив давлений
        double[] PLast = new double[3];

        double[] dPdt = new double[4];
        double[] Ts = new double[4];
        double[] T = new double[4];
        double[] dTdt = new double[4];
        double[] dQdt = new double[4];
        double[] V = new double[4];
        double[] D = new double[3];
        double[] Mu = new double[3];
        double[] G = new double[3];
        double[] TPmax = new double[3];

        double[,] PRes = new double[4, countOfArr];
        double[,] TRes = new double[4, countOfArr];
        double[,] GRes = new double[3, countOfArr];

        double DT = new double();
        double Alpha = new double();

        int pY = new int(); //Переменная для сдвига окон графиков

        //Создание объектов класса рассчёта системы
        PneumoCalc Calc = new PneumoCalc();

        public MainForm()
        {
            InitializeComponent(); //Инициализация компонентов на форме окна Windows

        }
        /// <summary>
        /// Перевод текста в элементы массива
        /// </summary>
        public void ReadData()
        {

            try
            {
                //Перевод данных в double
                P[1] = double.Parse(textP1.Text) * Math.Pow(10,5);
                P[2] = double.Parse(textP2.Text) * Math.Pow(10, 5);
                P[3] = double.Parse(textP3.Text) * Math.Pow(10, 5);

                T[1] = double.Parse(textT1.Text);
                T[2] = double.Parse(textT2.Text);
                T[3] = double.Parse(textT3.Text);

                Ts[1] = double.Parse(textT1.Text);
                Ts[2] = double.Parse(textT2.Text);
                Ts[3] = double.Parse(textT3.Text);

                V[1] = double.Parse(textV1.Text) * Math.Pow(10, -3);
                V[2] = double.Parse(textV2.Text) * Math.Pow(10, -3);
                V[3] = double.Parse(textV3.Text) * Math.Pow(10, -3);

                D[1] = double.Parse(textD12.Text) * Math.Pow(10, -3);
                D[2] = double.Parse(textD13.Text) * Math.Pow(10, -3);

                Mu[1] = double.Parse(textMu12.Text);
                Mu[2] = double.Parse(textMu13.Text);

                DT = double.Parse(textStep.Text)/1000;

                //Определение модели расчёта (с теплообменом или без)
                if (radioTerm.Checked == true)
                    Alpha = double.Parse(textAlpha.Text);
                else Alpha = 0;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка ввода!");
            }

        }

        /// <summary>
        /// Нажатие на кнопку расчёт
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bCalc_Click(object sender, EventArgs e)
        {
            //Удаление лишних процессов Excel
            try
            {
                foreach (Process proc in Process.GetProcessesByName("Microsoft Excel"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }

            //Создание нового документа Excel
            Excel myExcel = new Excel();
            myExcel.NewDocument();

            //Создание объектов класса Графическое окно, их настройка и инициализация
            ChartForm PChart = new ChartForm
            {
                Location = new Point(this.Location.X, this.Location.Y + 450 + pY),
                Text = "График P (" + (pY / 50).ToString() + ")"
            };

            PChart.Show();
            PChart.chart.ChartAreas[0].AxisY.Title = "p, кПа";
            PChart.chart.ChartAreas[0].AxisY.Interval = 200;

            ChartForm TChart = new ChartForm
            {
                StartPosition = FormStartPosition.Manual,
                Location = new Point(PChart.Location.X + 650, PChart.Location.Y),
                Text = "График T (" + (pY / 50).ToString() + ")"
            };

            TChart.Show();
            TChart.chart.ChartAreas[0].AxisY.Title = "T, K";
            TChart.chart.ChartAreas[0].AxisY.Interval = 25;

            ChartForm GChart = new ChartForm
            {
                StartPosition = FormStartPosition.Manual,
                Location = new Point(TChart.Location.X + 650, PChart.Location.Y),
                Text = "График G (" + (pY / 50).ToString() + ")"
            };

            GChart.Show();
            GChart.chart.ChartAreas[0].AxisY.Title = "G, г/c";
            GChart.chart.ChartAreas[0].AxisY.Interval = 0.10;

            //Вывод по массовой концентрации
            ChartForm MChart = new ChartForm
            {
                StartPosition = FormStartPosition.Manual,
                Location = new Point(GChart.Location.X, PChart.Location.Y - 440),
                Text = "График g (" + (pY / 50).ToString() + ")"
            };

            MChart.Show();
            MChart.chart.ChartAreas[0].AxisY.Title = "g";
            MChart.chart.ChartAreas[0].AxisY.Interval = 0.1;

            pY = pY + 50; //Сдвиг для новых окон (при следующем нажатии на кнопку)
            
            //Настройка графиков
            for (int i = 0; i<3; i++)
            {
                PChart.chart.Series.Add("P"+(i+1).ToString());
                PChart.chart.Series[i].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                PChart.chart.Series[i].BorderWidth = 3;

                TChart.chart.Series.Add("T" + (i + 1).ToString());
                TChart.chart.Series[i].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                TChart.chart.Series[i].BorderWidth = 3;

            }

            for (int i = 0; i < 2; i++)
            {
                GChart.chart.Series.Add("G" + (i + 1).ToString());
                GChart.chart.Series[i].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                GChart.chart.Series[i].BorderWidth = 3;

                MChart.chart.Series.Add("g" + (i + 1).ToString());
                MChart.chart.Series[i].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                MChart.chart.Series[i].BorderWidth = 3;
            }

            //Получение данных газовой постоянной и показателя адиабаты для воздуха и гелия
            double RAir = Calc.RAir;
            double KAir = Calc.KAir;

            double RHe = Calc.RHe;
            double KHe = Calc.KHe;

            //Инициализация переменных для анализа max/min
            double Pmax = new double();
            double Pmin = new double();

            //Чтение данных с окна формы
            ReadData();

            double[] F = new double[3]; //Эффективное сечение
            
            //Определение Эффективных сечений
            for(int i=1; i<3; i++)
            {
                F[i] = Calc.Section(Mu[i], D[i]);
            }

            double[] S = new double[4]; //Площадь поверхности

            for (int i = 1; i < 4; i++)
            {
                S[i] = Calc.Surface(V[i]);
            }

            int n = 0;
            double tt = 0;
            int gt = 1000; //шаг вывода на графики
            string lbl = ""; //строка вывода

            //Вывод первой строчки заголовка в Excel
            myExcel.SetValue("A1", "t,c");

            myExcel.SetValue("B1", "P1,Па");
            myExcel.SetValue("C1", "P2,Па");
            myExcel.SetValue("D1", "P3,Па");


            myExcel.SetValue("F1", "T1,К");
            myExcel.SetValue("G1", "T2,К");
            myExcel.SetValue("H1", "T3,К");


            myExcel.SetValue("J1", "G12,г/c");
            myExcel.SetValue("K1", "G13,г/c");

            myExcel.SetValue("M1", "gHe");
            myExcel.SetValue("N1", "gAir");

            //Определение начальной массы воздуха в сосуде смеси (Согласно ТЗ, там имелся воздух)
            double MAir = P[1] * V[1] / (RAir * T[1]);
            
            //Инициализация переменной массы по гелию
            double MHe = 0;

            //Инициализация массовых концентраций
            double [] gHe = new double [countOfArr];
            double[] gAir = new double[countOfArr];

            //Массовая концентрация в начальный момент времени
            gAir [0] = 1;
            gHe [0] = 0;

            //Инициализация показателя адиабаты и удельной газовой постоянной для смеси
            double K = 0;
            double R = 0;

            do //ЦИКЛ
            {
                
                //Определение температур исходящих потоков
                if (P[1] > P[2]) TPmax[1] = T[1];
                else TPmax[1] = T[2];

                if (P[3] > P[1]) TPmax[2] = T[3];
                else TPmax[2] = T[1];


                //Определение расходов
                G[1] = Calc.Flow(F[1], P[2], P[1], TPmax[1], Gases.He);

                G[2] = Calc.Flow(F[2], P[1], P[3], TPmax[2], Gases.Air);

                //Определение масс после натекания за промежуток времени DT
                MAir = MAir - G[2] * DT; //Знак "минус" связан с тем, что расход имеет знак, и для воздуха он отрицательный (газ течёт против оси OX)
                MHe = MHe + G[1] * DT;

                //Определение Массовых концентраций
                gAir[n+1] = MAir /(MAir + MHe);
                gHe[n+1] = MHe / (MAir + MHe);

                //Определение показателя адиабаты и удельной газовой постоянной
                K = KHe * gHe[n + 1] + KAir * (1 - gHe[n + 1]);
                R = RHe * gHe[n + 1] + RAir * (1 - gHe[n + 1]);

                //Определение тепловых потоков
                for (int i =1; i<4; i++)
                {
                    dQdt[i] = Alpha * S[i] * (Ts[i] - T[i]);
                }

                //Вывод текущих данных в матрицу
                for (int i = 0; i < 2; i++)
                {
                    PRes[i,n] = P[i+1];
                    TRes[i,n] = T[i+1];
                    GRes[i,n] = G[i+1];
                }
                PRes[2, n] = P[2 + 1];
                TRes[2, n] = T[2 + 1];

                //БЛОК ВЫВОДА
                if ((n % gt == 0) | (n == 0))
                {
                    //Вывод данных в Excel
                    lbl = ((int)(n / gt) + 2).ToString();//переменная номера ячейки excel

                    myExcel.SetValue("A" + lbl, (Math.Round(tt,3)).ToString());

                    myExcel.SetValue("B" + lbl, (Math.Round(PRes[0, n])).ToString());
                    myExcel.SetValue("C" + lbl, (Math.Round(PRes[1, n])).ToString());
                    myExcel.SetValue("D" + lbl, (Math.Round(PRes[2, n])).ToString());

                    myExcel.SetValue("F" + lbl, (Math.Round(TRes[0, n],1)).ToString());
                    myExcel.SetValue("G" + lbl, (Math.Round(TRes[1, n],1)).ToString());
                    myExcel.SetValue("H" + lbl, (Math.Round(TRes[2, n],1)).ToString());

                    myExcel.SetValue("J" + lbl, (Math.Round(GRes[0, n] * 1000, 2)).ToString());
                    myExcel.SetValue("K" + lbl, (Math.Round(GRes[1, n] * 1000, 2)).ToString());

                    myExcel.SetValue("M" + lbl, (Math.Round(gHe[n], 3)).ToString());
                    myExcel.SetValue("N" + lbl, (Math.Round(gAir[n], 3)).ToString());


                    //Вывод графиков концентраций
                    MChart.chart.Series[0].Points.AddXY(Math.Round(tt, 3), Math.Round(gHe[n], 3));
                    MChart.chart.Series[1].Points.AddXY(Math.Round(tt, 3), Math.Round(gAir[n], 3));

                    //Вывод данных на график
                    for (int i = 0; i < 3; i++)
                    {
                        PChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), PRes[i, n] / 1000);
                        TChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), TRes[i, n]);
                    }
                    for (int i = 0; i < 2; i++)
                    {
                        GChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), GRes[i, n] * 1000);
                    }

                }

                n = n+1;
                tt = tt + DT;

                //Определение изменений давления и температуры
                dPdt[2] = -1 * KHe * RHe / V[2] * TPmax[1] * G[1] + KHe / V[2] * (KHe - 1)* dQdt[2];
                dTdt[2] = T[2] / (P[2] * V[2]) * (V[2] * dPdt[2] - (-1)* RHe * T[2]*G[1] + (KHe - 1) * dQdt[2]);

                dPdt[1] = K * R / V[1] * (TPmax[1] * G[1]- TPmax[2]* G[2]) + K / V[1] * (K - 1) * dQdt[1];
                dTdt[1] = T[1] / (P[1] * V[1]) * (V[1] * dPdt[1] - R * T[1] * (G[1] - G[2]) + (K - 1) * dQdt[1]);

                dPdt[3] = KAir * RAir / V[3] * (TPmax[2] * G[2]) + KAir / V[3] * (KAir - 1) * dQdt[3];
                dTdt[3] = T[3] / (P[3] * V[3]) * (V[3] * dPdt[3] - RAir * T[3] * G[2] + (KAir - 1) * dQdt[3]);

                PCheck.Clear(); //Очистка динамического массива
                
                //Пересчёт новых давлений
                for (int i = 1; i < 4; i++)
                {

                    PLast[i - 1] = P[i];

                    P[i] = P[i] + dPdt[i] * DT;
                    T[i] = T[i] + dTdt[i] * DT;

                    if (P[i] != PLast[i - 1]) //Если давление изменяется, то добавить к сравнению
                    PCheck.Add(P[i]);

                }

                //Определение максимального и минимального давлений, если есть, что сравнивать
                if (PCheck.LongCount() == 0)
                {
                    MessageBox.Show("Давления не изменятся!!!");
                    break;
                }
                else
                {
                    Pmax = PCheck.Max();
                    Pmin = PCheck.Min();
                }

            } while (Pmin<= 0.95*Pmax); //Условие окончания счёта - разница между max и min не больше 5%

            lbl = ((int)(n / gt) + 3).ToString();
            n--;

            //Вывод последних значений следующих величин: расходов, давлений, температур и массовых концентраций
            myExcel.SetValue("A" + lbl, (Math.Round(tt, 3)).ToString());

            myExcel.SetValue("B" + lbl, (Math.Round(PRes[0, n])).ToString());
            myExcel.SetValue("C" + lbl, (Math.Round(PRes[1, n])).ToString());
            myExcel.SetValue("D" + lbl, (Math.Round(PRes[2, n])).ToString());

            myExcel.SetValue("F" + lbl, (Math.Round(TRes[0, n], 1)).ToString());
            myExcel.SetValue("G" + lbl, (Math.Round(TRes[1, n], 1)).ToString());
            myExcel.SetValue("H" + lbl, (Math.Round(TRes[2, n], 1)).ToString());

            myExcel.SetValue("J" + lbl, (Math.Round(GRes[0, n] * 1000, 2)).ToString());
            myExcel.SetValue("K" + lbl, (Math.Round(GRes[1, n] * 1000, 2)).ToString());

            myExcel.SetValue("M" + lbl, (Math.Round(gHe[n], 3)).ToString());
            myExcel.SetValue("N" + lbl, (Math.Round(gAir[n], 3)).ToString());


            //Вывод данных на график
            for (int i = 0; i < 3; i++)
            {
                PChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), PRes[i, n] / 1000);
                TChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), TRes[i, n]);
            }
            for (int i = 0; i < 2; i++)
            {
                GChart.chart.Series[i].Points.AddXY(Math.Round(tt, 3), GRes[i, n] * 1000);
            }

            //Вывод графиков концентраций
            MChart.chart.Series[0].Points.AddXY(Math.Round(tt, 3), Math.Round(gHe[n + 1], 3));
            MChart.chart.Series[1].Points.AddXY(Math.Round(tt, 3), Math.Round(gAir[n + 1], 3));

            //Развёртывание окна excel
            myExcel.Visible = true;

        }

    }
    
}
