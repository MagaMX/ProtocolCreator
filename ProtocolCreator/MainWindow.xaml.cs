using DocumentFormat.OpenXml.Drawing;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ProtocolCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? filePath, liquid, sample;
        private int? pressure;
        private double? interval;
        private Storyboard BlinkingStoryboard;
        private WordFileWriter wordFileWriter;

        public MainWindow()
        {
            InitializeComponent();
            LoadWindow();
        }

        private void LoadWindow()
        {
            CreateBlinkingAnimation();
            tbWriteToProtocol.Visibility = Visibility.Hidden;
            tbSuccessMsg.Visibility = Visibility.Hidden;
            BlinkingEllipse.Visibility = Visibility.Hidden;
            btnWriteMode.Visibility = Visibility.Hidden;
            btnStopWrite.Visibility = Visibility.Hidden;
        }

        private void Button_Start(object sender, RoutedEventArgs e)
        {
            tbWriteToProtocol.Visibility = Visibility.Visible;
          
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word file (*.docx)|*.docx";
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (saveFileDialog.ShowDialog() == true)
            {
                filePath = saveFileDialog.FileName;
            }

            try
            {
                liquid = textBox_Liquid.Text;
                sample = textBox_Sample.Text;
                pressure = int.Parse(textBox_Pressure.Text);
                interval = double.Parse(textBox_Interval.Text, System.Globalization.CultureInfo.InvariantCulture);

                if (!string.IsNullOrEmpty(filePath) && !string.IsNullOrEmpty(liquid) && !string.IsNullOrEmpty(sample) && pressure.HasValue && interval.HasValue)
                {
                    textBox_Liquid.IsEnabled = false;
                    textBox_Sample.IsEnabled = false;
                    textBox_Pressure.IsEnabled = false;
                    textBox_Interval.IsEnabled = false;

                    TemplateCreator templateCreator = new(filePath, liquid, sample, pressure.Value, interval.Value);
                    templateCreator.Create();

                    wordFileWriter = new(filePath, interval.Value);
                }

                // Делаем индикатор видимым
                BlinkingEllipse.Visibility = Visibility.Visible;

                // Проверяем, была ли создана анимация и запускаем её
                if (BlinkingStoryboard != null)
                {
                    BlinkingStoryboard.Begin();
                }

                buttonStart.IsEnabled = false;
                btnWriteMode.Visibility = Visibility.Visible;
                btnStopWrite.Visibility = Visibility.Visible;
            }
            catch(Exception)
            {
                MessageBox.Show("Заполните все поля!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnWriteMode_Click(object sender, RoutedEventArgs e)
        {
            wordFileWriter.RecordMode();
            if(wordFileWriter.IsOk())
            {
                tbSuccessMsg.Visibility = Visibility.Visible;
                ShowResultMessage(true);
            }
            else
            {
                tbSuccessMsg.Visibility = Visibility.Visible;
                MessageBox.Show("Ошибка при работе с файлом");
                ShowResultMessage(false);
            }
        }

        private void btnStopWrite_Click(object sender, RoutedEventArgs e)
        {
            tbWriteToProtocol.Visibility = Visibility.Hidden;
            tbSuccessMsg.Visibility = Visibility.Hidden;
            BlinkingEllipse.Visibility = Visibility.Hidden;
            btnWriteMode.Visibility = Visibility.Hidden;
            btnStopWrite.Visibility = Visibility.Hidden;

            textBox_Liquid.IsEnabled = true;
            textBox_Sample.IsEnabled = true;
            textBox_Pressure.IsEnabled = true;
            textBox_Interval.IsEnabled = true;
            buttonStart.IsEnabled = true;

            wordFileWriter.RecordCHF();
            MessageBox.Show("Запись остановлена. Все изменения сохранены в файл протокола!", "Остановка записи", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CreateBlinkingAnimation()
        {
            // Создание анимации
            ColorAnimation colorAnimation = new ColorAnimation
            {
                From = Colors.Green,
                To = Colors.Transparent,
                Duration = new Duration(System.TimeSpan.FromSeconds(0.5)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever
            };

            // Создание Storyboard и привязка анимации
            BlinkingStoryboard = new Storyboard();
            Storyboard.SetTarget(colorAnimation, BlinkingEllipse);
            Storyboard.SetTargetProperty(colorAnimation, new PropertyPath("(Ellipse.Fill).(SolidColorBrush.Color)"));

            BlinkingStoryboard.Children.Add(colorAnimation);
        }

        // Асинхронный метод для показа сообщения на несколько секунд
        private async void ShowResultMessage(bool boolean)
        {
            if(boolean)
            {
                // Устанавливаем текст в TextBlock
                tbSuccessMsg.Text = "Запись режима успешна!";

                // Делаем задержку на 2 секунды
                await Task.Delay(2000);

                // Очищаем текст после задержки
                tbSuccessMsg.Text = "";
            }

            else
            {
                // Устанавливаем текст в TextBlock
                tbSuccessMsg.Text = "Записать режим не удалось!";

                // Делаем задержку на 2 секунды
                await Task.Delay(2000);

                // Очищаем текст после задержки
                tbSuccessMsg.Text = "";
            }

        }
    }
}