using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Bank
{
    /// <summary>
    /// Логика взаимодействия для AnalizeVklad.xaml
    /// </summary>
    public partial class AnalizeVklad : Window
    {
        int sum;
        int time;
        public AnalizeVklad(int sum,int time)
        {
            InitializeComponent();
            this.sum = sum;
            this.time = time;
        }

        private void btn_Create_Click(object sender, RoutedEventArgs e)
        {
            //Код антуха
            UIElement element = Screen as UIElement;
            Uri path = new Uri(@"C:\Users\Admin\source\repos\BankRevision3-main\BankRevision2-main\Bank\screenshot.png");
            CaptureScreen(element, path);
        }
        public void CaptureScreen(UIElement source, Uri destination)
        {
            try
            {
                double Height, renderHeight, Width, renderWidth;

                Height = renderHeight = source.RenderSize.Height;
                Width = renderWidth = source.RenderSize.Width;


                RenderTargetBitmap renderTarget = new RenderTargetBitmap((int)renderWidth, (int)renderHeight, 96, 96, PixelFormats.Pbgra32);

                VisualBrush visualBrush = new VisualBrush(source);
                //Код антуха
                DrawingVisual drawingVisual = new DrawingVisual();
                using (DrawingContext drawingContext = drawingVisual.RenderOpen())
                {

                    drawingContext.DrawRectangle(visualBrush, null, new Rect(new Point(0, 0), new Point(Width, Height)));
                }

                renderTarget.Render(drawingVisual);


                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderTarget));
                using (FileStream stream = new FileStream(destination.LocalPath, FileMode.Create, FileAccess.Write))
                {
                    encoder.Save(stream);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void btn_openStab_Click(object sender, RoutedEventArgs e)
        {
            string percent = "8";
            authorization form1 = new authorization(percent,sum,time);
            form1.Show();
            this.Close();
        }

        private void btn_OpenOpt_Click(object sender, RoutedEventArgs e)
        {
            string percent = "5";
            authorization form1 = new authorization(percent, sum, time);
            form1.Show();
            this.Close();
        }

        private void btn_openStandart_Click(object sender, RoutedEventArgs e)
        {
            string percent = "6";
            authorization form1 = new authorization(percent, sum, time);
            form1.Show();
            this.Close();
            
        }
    }
}
