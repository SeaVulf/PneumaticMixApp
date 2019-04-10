using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PnumSysCalc
{
    public class PneumoCalc
    {
        //Газовая постоянная воздуха, показатель адиабаты и критическое отношение давлений
        public readonly double R = 287.2;
        public readonly double K = 1.41;
        public readonly double BK = 0.528;

        /// <summary>
        /// Определение эффективной площади поперечного сечения условного сопротивления
        /// </summary>
        /// <param name="Mu">Коэффициент расхода</param>
        /// <param name="D">Диаметр условного сопротивления</param>
        /// <returns></returns>
        public double Section(double Mu, double D)
        {
            return Mu * Math.PI * D * D / 4;
        }

        public double Surface(double V)
        {
            return Math.PI*Math.Pow(6*V/Math.PI,2/3);
        }

        /// <summary>
        /// Определение направления расхода
        /// </summary>
        /// <param name="pL">Давление левого сосуда</param>
        /// <param name="pR">Давление правого сосуда</param>
        /// <returns></returns>
        public int FlowDirection(double pL, double pR)
        {
            if (pL > pR) return 1;
            else return -1;
        }
        /// <summary>
        /// Определение расхода
        /// </summary>
        /// <param name="F">Площадь эффективного поперечного сечения условного сопротивления</param>
        /// <param name="pL">Давление "левого" сосуда</param>
        /// <param name="pR">Давление "правого" сосуда</param>
        /// <param name="TpA">Температура сосуда с наибольшим давлением</param>
        /// <returns></returns>
        public double Flow(double F, double pL, double pR, double TpA)
        {
            try
            {
                //Определение максимальных давлений
                double pA = Math.Max(pL, pR);
                double pB = Math.Min(pL, pR);

                //Расчёт расхода с учётом направления истечения
                double G = FlowDirection(pL, pR) * F * pA * Math.Sqrt(2 * K / ((K - 1) * R * TpA)) * Math.Sqrt(Math.Pow(Math.Max(pB / pA, BK), 2 / K) - Math.Pow(Math.Max(pB / pA, BK), (K + 1) / K));

                return G;

            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка Определения расхода!");
                return 0;
            }


        }
    }
}
