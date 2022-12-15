using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.TextFormatting;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;


namespace Bank
{
    /// <summary>
    /// Логика взаимодействия для authorization.xaml
    /// </summary>
    public partial class authorization : Window
    {
        string percent;
        int sum;
        int time;

        public authorization(string percent,int sum,int time)
        {
            InitializeComponent();
            this.percent = percent;
            this.sum = sum;
            this.time = time;
            
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //txb_login.Clear();

        }
       

        private void txb_login_MouseDown(object sender, MouseButtonEventArgs e)
        {
           // txb_login.Clear();
            //txb_login.Focus();
            //txb_login.Text = "";

        }

        private void txb_login_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //txb_login.Clear();
        }
       
        string id1;
        string NumberAccount1;
        string UserID1;
        string DateOpen1;
        string Balance1;
        string TypeID1;        
        string Login2;
        string Password2;
        string Name1;
        string surname1;
        string Patronymic1;
        string series1;
        string Number1;
        string Phone1;
        string Adress1;
        string E_Mail1;
        string DateOfIssue1;
        string Issueed1;
        string DateOfBirth1;
        string PlaceOfBirth1;
        string IDContract1;
        private void btn_login_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = @"Data Source=DESKTOP-4JK1FOK\SQLEXPRESS;Initial Catalog=bank;Integrated Security=True";
            string login = txb_login.Text;
            string password = txt_pass.Text;
            string sqlExpression = "SELECT * FROM [dbo].[User] where Login =" + "'"+login+"'"+  "and "+ "Password = "+"'" + password + "'";
            //txt_pass.Text = sqlExpression;
            
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression, connection);
                SqlDataReader reader = command.ExecuteReader();
                int c = 0;
                if (reader.HasRows) // если есть данные
                {
                    

                    while (reader.Read()) // построчно считываем данные
                    {
                        if (c==1)
                        {
                            break;
                        }
                        object id = reader.GetValue(0);
                        object Login1 = reader.GetValue(1);
                        object Password1 = reader.GetValue(2);
                        object Name = reader.GetValue(3);
                        object surname = reader.GetValue(4);
                        object Patronymic = reader.GetValue(5);
                        object series = reader.GetValue(6);
                        object Number = reader.GetValue(7);
                        object Phone = reader.GetValue(8);
                        object Adress = reader.GetValue(9);
                        object E_Mail = reader.GetValue(10);
                        object DateOfIssue   = reader.GetValue(11);
                        object Issueed = reader.GetValue(12);
                        object DateOfBirth = reader.GetValue(13);
                        object PlaceOfBirth = reader.GetValue(14);
                        id1 = Convert.ToString(id);                       
                        Name1 = Convert.ToString(Name);
                        surname1 = Convert.ToString(surname);
                        Patronymic1 = Convert.ToString(Patronymic);
                        series1 = Convert.ToString(series);
                        Number1 = Convert.ToString(Number);
                        Phone1 = Convert.ToString(Phone) ;
                        Adress1 = Convert.ToString(Adress);
                        E_Mail1 = Convert.ToString(E_Mail);
                        DateOfIssue1 = Convert.ToString(DateOfIssue);
                        Issueed1 = Convert.ToString(Issueed);
                        DateOfBirth1 = Convert.ToString(DateOfBirth);
                        PlaceOfBirth1 = Convert.ToString(PlaceOfBirth);
                        c = c + 1;                        
                    }
                }


                reader.Close();
            }
                          
            string sqlExpression1 = "SELECT * FROM [dbo].[BankAccount] where UserID =" + "'" + id1 + "'";            
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression1, connection);
                SqlDataReader reader = command.ExecuteReader();
                int c = 0;
                if (reader.HasRows) // если есть данные
                {


                    while (reader.Read()) // построчно считываем данные
                    {
                        if (c == 1)
                        {
                            break;
                        }
                        object NumberAccount = reader.GetValue(0);
                        object UserID = reader.GetValue(1);
                        object DateOpen = reader.GetValue(2);
                        object Balance = reader.GetValue(3);
                        object TypeID = reader.GetValue(4);
                        NumberAccount1 = Convert.ToString(NumberAccount);
                        UserID1 = Convert.ToString(UserID) ;
                        DateOpen1 = Convert.ToString(DateOpen);
                        Balance1 = Convert.ToString(Balance);
                        TypeID1 = Convert.ToString(TypeID);
                    }
                }


                reader.Close();
            }
            string sqlExpression2 = "SELECT * FROM [dbo].[Contract] where UserID =" + "'" + id1 + "'";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression2, connection);
                SqlDataReader reader = command.ExecuteReader();
                int c = 0;
                if (reader.HasRows) // если есть данные
                {


                    while (reader.Read()) // построчно считываем данные
                    {
                        if (c == 1)
                        {
                            break;
                        }
                        object IDContract = reader.GetValue(0); //у этого бедного егора например нет контракта.Так что создадим ему сами :D
                        IDContract1 = Convert.ToString(IDContract);
                        
                        
                    }
                }
                if (IDContract1 == null)
                {
                    Random r = new Random();
                    IDContract1 = Convert.ToString(r.Next(25, 100));
                }

                reader.Close();
            }
            string sum1=Convert.ToString(sum);
            string time1 = Convert.ToString(time);
            string TemplateFileName = @"C:\Users\Admin\source\repos\BankRevision3-main\BankRevision2-main\Bank\Shablon_dogovora.docx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var dt = DateTime.Now.AddMonths(time).Date;
            string Issueed2 = Issueed1 + " "+ DateOfIssue1;
            try
            {
                string FIO = Name1 + " "+surname1 +" "+ Patronymic1;
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWordStub("<NUMDOG>", IDContract1, wordDocument);
                ReplaceWordStub("<DAY>", DateTime.Today.ToString("dd"), wordDocument);
                ReplaceWordStub("<MONTHS>", DateTime.Today.ToString("MM"), wordDocument);
                ReplaceWordStub("<YEAR>", DateTime.Today.ToString("yy"), wordDocument);
                ReplaceWordStub("<FIOVKLAD>", FIO, wordDocument);
                ReplaceWordStub("<SUMVKLAD>", sum1, wordDocument);
                ReplaceWordStub("<TIMEVKLD>", time1, wordDocument);
                ReplaceWordStub("<DATEENDVKLAD>", dt.ToString(), wordDocument);
                ReplaceWordStub("<PERCENTVKLAD>", percent, wordDocument);
                ReplaceWordStub("<NUMBERACC>", NumberAccount1, wordDocument);
                ReplaceWordStub("<ADESSREG>", Adress1, wordDocument);
                ReplaceWordStub("<EMAIL>", E_Mail1, wordDocument);
                ReplaceWordStub("<NM>", Number1, wordDocument);
                ReplaceWordStub("<ISSUE>", Issueed2, wordDocument);
                ReplaceWordStub("<DATEBIRTH>", DateOfBirth1, wordDocument);
                ReplaceWordStub("<PLACEBIRTH>", PlaceOfBirth1, wordDocument);
                ReplaceWordStub("<SERIES>", series1, wordDocument);
                ReplaceWordStub("<FIOVKLAD1>", FIO, wordDocument);
                SaveFileDialog sfd = new SaveFileDialog();
                Nullable<bool> result = sfd.ShowDialog();
                      
                if (result == true)
                {
                    DateTime localDate = DateTime.Now;
                    wordDocument.SaveAs(@"C:\Users\Admin\source\repos\BankRevision3-main\BankRevision2-main\Bank\dogovor.docx");
                    MessageBox.Show("Договор сохранен");
                    this.Close();
                }
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }
        private void ReplaceWordStub(string stubToReplace,string text, Word.Document worddocument)
        {
            var range = worddocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void txt_pass_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
 }

//11198451
//CiYKA519tAMlqktBk7