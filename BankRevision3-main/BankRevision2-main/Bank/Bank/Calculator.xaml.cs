using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Bank
{
    /// <summary>
    /// Логика взаимодействия для Calculator.xaml
    /// </summary>
    public partial class Calculator : Window
    {
        int C;
        int T;
        public double itog;
        double itog2;
        double itog3;
        public Calculator()
        {
            InitializeComponent();
           
        }
        


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AnalizeVklad form1 = new AnalizeVklad(C,T);
            form1.Show();
            this.Close();
            form1.txb_sumStab.Text = txt_moneyout.Text;
            form1.txb_SumOpt.Text = txt_moneyoptim.Text;
            form1.txb_SumStandart.Text = txt_moneystandart.Text;
            double EndSum = C+itog;
            double EndSum2 = C + itog2;
            double EndSum3 = C + itog3;
            form1.txb_endSumStab.Text = EndSum.ToString() +" "+"Руб.";
            form1.txb_EndSumOpt.Text = EndSum2.ToString() + " " + "Руб.";
            form1.txb_EndSumStandart.Text = EndSum3.ToString() + " " + "Руб.";
        }
       
        private void slValue_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            try
            {
               
                C = Convert.ToInt32(txb_money.Text);
                T = Convert.ToInt32(txt_mouth.Text);
                double Stab = 0.08 / 12;
                Stab = Stab + 1;
                double Dv = C * ((Math.Pow(Stab, T)));
                //txt_moneyout.Text = Dv.ToString();
                double inter = Dv - C;
                itog = Math.Round(inter);
                txt_moneyout.Text = itog.ToString() + " "+ "Руб.";
                double Opt = 0.05 / 12;
                Opt = Opt + 1; 
                double Dv2 = C * ((Math.Pow(Opt, T)));
                //txt_moneyout.Text = Dv.ToString();
                double inter2 = Dv2 - C;
                itog2 = Math.Round(inter2);
                txt_moneyoptim.Text = itog2.ToString() + " " + "Руб.";
                double Standart = 0.06 / 12;
                Standart = Standart + 1;
                double Dv3 = C * ((Math.Pow(Standart, T)));
                //txt_moneyout.Text = Dv.ToString();
                double inter3 = Dv3 - C;
                itog3 = Math.Round(inter3);
                txt_moneystandart.Text = itog3.ToString() + " " + "Руб.";



            }
            catch
            {

            }


        }
    }
}
