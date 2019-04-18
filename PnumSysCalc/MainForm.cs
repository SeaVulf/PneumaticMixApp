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
        double[] P = new double[4];
        List<double> PCheck = new List<double>(); //Динамический массив
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

        double[,] PRes = new double[4, 10000000];
        double[,] TRes = new double[4, 10000000];
        double[,] GRes = new double[3, 10000000];

        double DT = new double();
        double Alpha = new double();

        bool flag = true;
        bool[] flagArr = new bool[3];

        int pY = new int();

        //Создание объектов классов для расчёта
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


            Excel myExcel = new Excel();
            myExcel.NewDocument();//Создание нового документа Excel

            //Создание объектов класса Графическое окно и их настройка
            ChartForm PChart = new ChartForm
            {
                Location = new Point(this.Location.X, this.Location.Y + 450 + pY),
                Text = "График P (" + (pY / 50).ToString() + ")"
            };

            //Инициализация графика
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

            pY = pY + 50;
            
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
            }

            //Получение данных газовой постоянной и показателя адиабаты
            double R = Calc.RAir;
            double KAir = Calc.KAir;
            double K = Calc.KAir;

            //Проверка работоспособности функции
            double cHeck = Calc.BKr(KAir);

            double Pmax = new double();
            double Pmin = new double();

            //Определение данных
            ReadData();

            double[] F = new double[4]; //Эффективное сечение
            
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

            do //ЦИКЛ
            {
                
                //Определение температур исходящих потоков
                if (P[1] > P[2]) TPmax[1] = T[1];
                else TPmax[1] = T[2];

                if (P[3] > P[1]) TPmax[2] = T[3];
                else TPmax[2] = T[1];


                //Определение расходов
                G[1] = Calc.Flow(F[1], P[2], P[1], TPmax[1]);

                G[2] = Calc.Flow(F[2], P[1], P[3], TPmax[2]);

//TODO - учитывать газы!!!

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

                if ((n % gt == 0) | (n == 0))  {
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

//TODO
                //Определение изменений давления и температуры
                dPdt[2] = -1 * K * R / V[2] * TPmax[1] * (G[1]+ G[4]) + K/ V[2] * (K - 1)* dQdt[2];
                dTdt[2] = T[2] / (P[2] * V[2]) * (V[2] * dPdt[2] - (-1)*R* T[2]*(G[1] + G[4]) + (K - 1) * dQdt[2]);

                dPdt[1] = K * R / V[1] * (TPmax[1] * (G[1] + G[4])- TPmax[2]* G[2]) + K / V[1] * (K - 1) * dQdt[1];
                dTdt[1] = T[1] / (P[1] * V[1]) * (V[1] * dPdt[1] - R * T[1] * (G[1] + G[4]- G[2]) + (K - 1) * dQdt[1]);

                dPdt[3] = K * R / V[3] * (TPmax[2] * G[2] - TPmax[3] * (G[3]+G[5])) + K / V[3] * (K - 1) * dQdt[3];
                dTdt[3] = T[3] / (P[3] * V[3]) * (V[3] * dPdt[3] - R * T[3] * (G[2] - (G[3] + G[5])) + (K - 1) * dQdt[3]);

                dPdt[4] = K * R / V[4] * (TPmax[3] * (G[3] + G[5])) + K / V[4] * (K - 1) * dQdt[4];
                dTdt[4] = T[4] / (P[4] * V[4]) * (V[4] * dPdt[4] - R * T[4] * (G[3] + G[5]) + (K - 1) * dQdt[4]);

                PCheck.Clear(); //Очистка динамического массива
                //Пересчёт новых давлений
                for (int i = 1; i < 5; i++)
                {
                    flagArr[i - 1] = true;

                    PLast[i - 1] = P[i];

                    P[i] = P[i] + dPdt[i] * DT;
                    T[i] = T[i] + dTdt[i] * DT;

                    if (P[i] != PLast[i - 1]) //Если давление изменяется, то добавить к сравнению
                    PCheck.Add(P[i]);

                    //Проверка расхождения
                    if (Math.Abs((P[i] - PLast[i - 1]) / PLast[i - 1]) < 0.0000001) { flagArr[i - 1] = false; }
                    else flagArr[i - 1] = true;
                }

                if (flagArr.Max() == false) flag = false;

                //Определение максимального и минимального давлений

                if (PCheck.LongCount() == 0) MessageBox.Show("Давления не изменятся!!!");
                else { 
                Pmax = PCheck.Max();
                Pmin = PCheck.Min();
                }

            } while ((Pmin<=0.95* Pmax)&(flag)); //Условие окончания счёта - разница между max и min не больше 5%

            lbl = ((int)(n / gt) + 3).ToString();
            n--;

            //Вывод последних значений следующих величин: расходов, давлений, температур
            myExcel.SetValue("A" + lbl, (Math.Round(tt, 3)).ToString());

            myExcel.SetValue("B" + lbl, (Math.Round(PRes[0, n])).ToString());
            myExcel.SetValue("C" + lbl, (Math.Round(PRes[1, n])).ToString());
            myExcel.SetValue("D" + lbl, (Math.Round(PRes[2, n])).ToString());

            myExcel.SetValue("F" + lbl, (Math.Round(TRes[0, n], 1)).ToString());
            myExcel.SetValue("G" + lbl, (Math.Round(TRes[1, n], 1)).ToString());
            myExcel.SetValue("H" + lbl, (Math.Round(TRes[2, n], 1)).ToString());

            myExcel.SetValue("J" + lbl, (Math.Round(GRes[0, n] * 1000, 2)).ToString());
            myExcel.SetValue("K" + lbl, (Math.Round(GRes[1, n] * 1000, 2)).ToString());

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
      
        //Развёртывание окна excel
        myExcel.Visible = true;


  //TODO
            //ПЕРЕТИРАНИЕ ВСЕГО
        }

    }
    
}
