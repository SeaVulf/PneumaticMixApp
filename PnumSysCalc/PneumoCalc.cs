using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PnumMixCalc
{
    //Перечисляемый тип воздух, гелий
    public enum Gases { Air, He, Mix };
    public class PneumoCalc
    {
        //Газовая постоянная и показатель адиабаты для воздуха и Не
        public readonly double RAir = 287.2;
        public readonly double KAir = 1.41;

        public readonly double RHe = 2077.4;
        public readonly double KHe = 1.66;

        //Инициализация переменных для хранения критического перепада давления гелия и воздуха
        double BKHe;
        double BKAir;

        //Конструктор класса, производящий рассчёт критических перепадов давления
        public PneumoCalc()
        {
            BKHe = BKr(KHe);
            BKAir = BKr(KAir);
        }

        /// <summary>
        /// Расчёт критического перепада давления
        /// </summary>
        /// <param name="K">Показатель адиабаты соответствующего газа</param>
        /// <returns></returns>
        public double BKr(double K)
        {
            return Math.Pow(2/(K+1),K/(K-1));
        }

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

        /// <summary>
        /// Расчёт площади поверхностей сосудов
        /// </summary>
        /// <param name="V"></param>
        /// <returns></returns>
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
        /// <param name="b">Род чистого газа</param>
        /// <returns></returns>
        public double Flow(double F, double pL, double pR, double TpA, Gases b)
        {
            try
            {
                //Определение максимального и минимального давлений
                double pA = Math.Max(pL, pR);
                double pB = Math.Min(pL, pR);

                //Инициализация переменных
                double BK;
                double K;
                double R;
                double G;

                //Определение характеристики газа в зависимости от его рода
                switch (b)
                {
                    case Gases.Air:
                        BK = BKAir;
                        K = KAir;
                        R = RAir;
                    break;

                    case Gases.He:
                        BK = BKHe;
                        K = KHe;
                        R = RHe;
                    break;

                    default:
                        BK = 0;
                        K = 0;
                        R = 0;
                    break;

                }

                //Расчёт возможного расхода (без учёта обратного клапана)
                G = FlowDirection(pL, pR) * F * pA * Math.Sqrt(2 * K / ((K - 1) * R * TpA)) * Math.Sqrt(Math.Pow(Math.Max(pB / pA, BK), 2 / K) - Math.Pow(Math.Max(pB / pA, BK), (K + 1) / K));

                //Обратный клапан (зануление расхода в случая обратного потока)
                switch (b)
                {
                    case Gases.Air:
                        if (G > 0)
                        {
                            G = 0;
                        }
                        break;

                    case Gases.He:
                        if (G < 0)
                        {
                            G = 0;
                        }
                        break;

                    default:
                        G = 0;
                        break;

                }

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
