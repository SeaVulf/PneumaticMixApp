using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PnumSysCalc
{
    public partial class MainForm : Form
    {
        double[] P = new double[5];
        List<double> PCheck = new List<double>(); //Динамический массив
        double[] PLast = new double[4];

        double[] dPdt = new double[5];
        double[] T = new double[5];
        double[] dTdt = new double[5];
        double[] V = new double[5];
        double[] D = new double[6];
        double[] Mu = new double[6];
        int[] A = new int[6];
        double[] G = new double[6];
        double[] TPmax = new double[4];

        double[,] PRes = new double[5, 10000000];
        double[,] TRes = new double[5, 10000000];
        double[,] GRes = new double[6, 10000000];

        double DT = new double();

        bool flag = true;
        bool[] flagArr = new bool[4];

        //Создание объектов классов для расчёта и для работы с Excel
        PneumoCalc Calc = new PneumoCalc();
        Excel myExcel = new Excel();

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
                P[4] = double.Parse(textP4.Text) * Math.Pow(10, 5);

                T[1] = double.Parse(textT1.Text);
                T[2] = double.Parse(textT2.Text);
                T[3] = double.Parse(textT3.Text);
                T[4] = double.Parse(textT4.Text);

                V[1] = double.Parse(textV1.Text) * Math.Pow(10, -3);
                V[2] = double.Parse(textV2.Text) * Math.Pow(10, -3);
                V[3] = double.Parse(textV3.Text) * Math.Pow(10, -3);
                V[4] = double.Parse(textV4.Text) * Math.Pow(10, -3);

                D[1] = double.Parse(textD12.Text) * Math.Pow(10, -3);
                D[2] = double.Parse(textD13.Text) * Math.Pow(10, -3);
                D[3] = double.Parse(textD34.Text) * Math.Pow(10, -3);
                D[4] = double.Parse(textD112.Text) * Math.Pow(10, -3);
                D[5] = double.Parse(textD334.Text) * Math.Pow(10, -3);

                Mu[1] = double.Parse(textMu12.Text);
                Mu[2] = double.Parse(textMu13.Text);
                Mu[3] = double.Parse(textMu34.Text);
                Mu[4] = double.Parse(textMu112.Text);
                Mu[5] = double.Parse(textMu334.Text);

                DT = double.Parse(textStep.Text)/1000;

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
            myExcel.NewDocument();//Создание нового документа Excel

            //Получение данных газовой постоянной и показателя адиабаты
            double R = Calc.R;
            double K = Calc.K;

            double Pmax = new double();
            double Pmin = new double();

            //Определение данных
            ReadData();

            double[] F = new double[6];
            
            //Определение Эффективных сечений
            for(int i=1; i<6; i++)
            {
                F[i] = Calc.Section(Mu[i], D[i]);
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
            myExcel.SetValue("E1", "P4,Па");

            myExcel.SetValue("F1", "T1,К");
            myExcel.SetValue("G1", "T2,К");
            myExcel.SetValue("H1", "T3,К");
            myExcel.SetValue("I1", "T4,К");

            myExcel.SetValue("J1", "G12,г/c");
            myExcel.SetValue("K1", "G13,г/c");
            myExcel.SetValue("L1", "G34,г/c");
            myExcel.SetValue("M1", "G112,г/c");
            myExcel.SetValue("N1", "G334,г/c");



            do //ЦИКЛ
            {
                flag = true;
                
                //Определение температур исходящих потоков
                if (P[1] > P[2]) TPmax[1] = T[1];
                else TPmax[1] = T[2];

                if (P[3] > P[1]) TPmax[2] = T[3];
                else TPmax[2] = T[1];

                if (P[4] > P[3]) TPmax[3] = T[4];
                else TPmax[3] = T[3];


                //Определение расходов
                G[1] = Calc.Flow(F[1], P[2], P[1], TPmax[1]);
                G[4] = Calc.Flow(F[4], P[2], P[1], TPmax[1]);

                G[2] = Calc.Flow(F[2], P[1], P[3], TPmax[2]);

                G[3] = Calc.Flow(F[3], P[3], P[4], TPmax[3]);
                G[5] = Calc.Flow(F[5], P[3], P[4], TPmax[3]);

                //Вывод текущих данных в матрицу
                for (int i = 0; i < 4; i++)
                {
                    PRes[i,n] = P[i+1];
                    TRes[i,n] = T[i+1];
                    GRes[i,n] = G[i+1];
                }
                GRes[4,n] = G[5];
                
                //Вывод данных в Excel
                if ((n % gt == 0) | (n == 0))  {

                    lbl = ((int)(n / gt) + 2).ToString();//переменная номера ячейки excel


                myExcel.SetValue("A" + lbl, (Math.Round(tt,3)).ToString());

                myExcel.SetValue("B" + lbl, (Math.Round(PRes[0, n])).ToString());
                myExcel.SetValue("C" + lbl, (Math.Round(PRes[1, n])).ToString());
                myExcel.SetValue("D" + lbl, (Math.Round(PRes[2, n])).ToString());
                myExcel.SetValue("E" + lbl, (Math.Round(PRes[3, n])).ToString());

                myExcel.SetValue("F" + lbl, (Math.Round(TRes[0, n],1)).ToString());
                myExcel.SetValue("G" + lbl, (Math.Round(TRes[1, n],1)).ToString());
                myExcel.SetValue("H" + lbl, (Math.Round(TRes[2, n],1)).ToString());
                myExcel.SetValue("I" + lbl, (Math.Round(TRes[3, n],1)).ToString());

                myExcel.SetValue("J" + lbl, (Math.Round(GRes[0, n] * 1000, 2)).ToString());
                myExcel.SetValue("K" + lbl, (Math.Round(GRes[1, n] * 1000, 2)).ToString());
                myExcel.SetValue("L" + lbl, (Math.Round(GRes[2, n] * 1000, 2)).ToString());
                myExcel.SetValue("M" + lbl, (Math.Round(GRes[3, n] * 1000, 2)).ToString());
                myExcel.SetValue("N" + lbl, (Math.Round(GRes[4, n] * 1000, 2)).ToString());

                }

                n = n+1;
                tt = tt + DT;

                //Определение изменений давления и температуры
                dPdt[2] = -1 * K * R / V[2] * TPmax[1] * (G[1]+ G[4]);
                dTdt[2] = T[2] / (P[2] * V[2]) * (V[2] * dPdt[2] - (-1)*R* T[2]*(G[1] + G[4]));

                dPdt[1] = K * R / V[1] * (TPmax[1] * (G[1] + G[4])- TPmax[2]* G[2]);
                dTdt[1] = T[1] / (P[1] * V[1]) * (V[1] * dPdt[1] - R * T[1] * (G[1] + G[4]- G[2]));

                dPdt[3] = K * R / V[3] * (TPmax[2] * G[2] - TPmax[3] * (G[3]+G[5]));
                dTdt[3] = T[3] / (P[3] * V[3]) * (V[3] * dPdt[3] - R * T[3] * (G[2] - (G[3] + G[5])));

                dPdt[4] = K * R / V[4] * (TPmax[3] * (G[3] + G[5]));
                dTdt[4] = T[4] / (P[4] * V[4]) * (V[4] * dPdt[4] - R * T[4] * (G[3] + G[5]));

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
            myExcel.SetValue("A" + lbl, (Math.Round(tt, 2)).ToString());

            myExcel.SetValue("B" + lbl, (Math.Round(PRes[0, n])).ToString());
            myExcel.SetValue("C" + lbl, (Math.Round(PRes[1, n])).ToString());
            myExcel.SetValue("D" + lbl, (Math.Round(PRes[2, n])).ToString());
            myExcel.SetValue("E" + lbl, (Math.Round(PRes[3, n])).ToString());

            myExcel.SetValue("F" + lbl, (Math.Round(TRes[0, n], 1)).ToString());
            myExcel.SetValue("G" + lbl, (Math.Round(TRes[1, n], 1)).ToString());
            myExcel.SetValue("H" + lbl, (Math.Round(TRes[2, n], 1)).ToString());
            myExcel.SetValue("I" + lbl, (Math.Round(TRes[3, n], 1)).ToString());

            myExcel.SetValue("J" + lbl, (Math.Round(GRes[0, n] * 1000, 2)).ToString());
            myExcel.SetValue("K" + lbl, (Math.Round(GRes[1, n] * 1000, 2)).ToString());
            myExcel.SetValue("L" + lbl, (Math.Round(GRes[2, n] * 1000, 2)).ToString());
            myExcel.SetValue("M" + lbl, (Math.Round(GRes[3, n] * 1000, 2)).ToString());
            myExcel.SetValue("N" + lbl, (Math.Round(GRes[4, n] * 1000, 2)).ToString());

            //Развёртывание окна excel, сохранение, закрытие
            myExcel.Visible = true;
           // myExcel.SaveDocument("Mine");
            myExcel.CloseDocument();

        }

       
    }
}
