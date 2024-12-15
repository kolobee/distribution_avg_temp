using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Numerics;
using Word = Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.ConstrainedExecution;
using System.Reflection;
using System.Reflection.Emit;
using LiveCharts;
using System.Windows.Markup;
using LiveCharts.Wpf;
using LiveCharts.Defaults;
using System.Xml.Linq;
using System.Windows.Controls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Drawing.Imaging;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;


namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        double alfa = 0;
        int zakon = 0;
        double M, D, Sigma;

        double h;
        int k;
        double max = int.MinValue;
        double min = int.MaxValue;

        double h1;

        const double Exp = 2.71828;
        const double pi = 3.14159265;

        List<double> p = new List<double>();
        List<double> np = new List<double>();
        List<double> n = new List<double>();
        List<double> nnp2 = new List<double>();
        List<double> nnp2np = new List<double>();

        static List<int> tmp = new List<int>();
        List<double> tmp1 = new List<double>();
        List<double> tmp2 = new List<double>();

        List<double> lyamda = new List<double>();

        List<double> LAPLAS_TABLE = new List<double>();
        List<double> X2_TABLE = new List<double>();                             // Список для значений таблицы X^2


        List<double> DATA = new List<double>();
        List<string> DATA1 = new List<string>();  // Список для хранения считанных из файла данных

        double arithmeticMean;                                                  // Среднее арифметическое 

        string filePath = string.Empty;

        public Form1()
        {
            InitializeComponent();

            string[] fileLines = File.ReadAllLines(@"../../LaplasTable.csv");   // Считываем все строки из файла таблицы Лапласа 
            foreach (string line in fileLines)                                  // Заносим строки в список 
            {
                string[] valueArray = line.Split(';');
                for (int i = 0; i < valueArray.Length; i++)
                {
                    LAPLAS_TABLE.Add(Convert.ToDouble(valueArray[i]));
                }
            }

            fileLines = File.ReadAllLines(@"../../x2.csv");                     // Делаем тоже самое для таблицы X^2 
            foreach (string line in fileLines)
            {
                string[] valueArray = line.Split(';');
                for (int i = 0; i < valueArray.Length; i++)
                {
                    X2_TABLE.Add(Convert.ToDouble(valueArray[i]));
                }
            }

            chart1.Legends.Clear();
        }
        // Загрузите CSV-файл в массив строк и столбцов.
        // Предположим, что могут быть пустые строки, но каждая строка имеет
        // столько же полей.
        private string[,] LoadCsv(string filename)
        {
            // Получить текст файла.
            string fileContent = textBox1.Text;

            // Разделение на строки.
            fileContent = fileContent.Replace('\n', '\r');
            string[] lines = fileContent.Split(new char[] { '\r' },
                StringSplitOptions.RemoveEmptyEntries);

            // Посмотрим, сколько строк и столбцов есть.
            int num_rows = lines.Length;
            int num_cols = lines[0].Split(';').Length;
                // Выделите массив данных.
                string[,] values = new string[num_rows, num_cols];

                // Загрузите массив.
                for (int r = 0; r < num_rows; r++)
                {
                    string[] line_r = lines[r].Split(';');
                    for (int c = 0; c < num_cols; c++)
                    {
                        values[r, c] = line_r[c];
                    }
                }

                // Возвращаем значения.
                return values;
            
        }
        private void CheckOnMisses()                                            // Функция проверки значений на промахи 
        {
            double calculatedIrvinValue;                                        // Расчетное значение критерия Ирвина 
            double tableIrvinValue;                                             // Табличное значение критерия Ирвина 
            double middleQuadDeviation;                                         // Выборочное среднеквадратичное отклонение 
            double suspectValue;                                                // Проверяемое значение из выборки 
            double previousValue;                                               // Значение предшествующее проверяемому 

            DATA.Sort();                                                        // Сортировка списка данных по возрастанию 

            while (true)                                                        // Проверяем, пока удается находить промахи 
            {
                suspectValue = DATA[DATA.Count() - 1];                          // В качестве проверяемого значения берем максимальное 
                previousValue = DATA[DATA.Count() - 2];

                double sum = 0;                                                 // Переменная для хранения промежуточных результатов при вычислении СКО 

                for (int i = 0; i < DATA.Count(); i++)
                {
                    sum += Math.Pow(DATA[i] - (DATA.Sum() / DATA.Count()), 2);
                }

                middleQuadDeviation = Math.Sqrt(sum / (DATA.Count() - 1));
                calculatedIrvinValue = (suspectValue - previousValue) / middleQuadDeviation;

                // В зависимости от уровня значимости вычисляем табличное значение критерия Ирвина 

                if (alfa == 1)                                     // Уровень значимости = 0.01 
                {
                    tableIrvinValue = -104.84 * Math.Pow(DATA.Count(), -3) + 257.04 * Math.Pow(DATA.Count(), -2.5) - 249.02 * Math.Pow(DATA.Count(), -2) +
                        114.69 * Math.Pow(DATA.Count(), -1.5) - 29.78 * Math.Pow(DATA.Count(), -1) + 6.217 * Math.Pow(DATA.Count(), -0.5) + 1.0504;
                }

                else if (alfa == 2)                                // Уровень значимости = 0.05 
                {
                    tableIrvinValue = -89.417 * Math.Pow(DATA.Count(), -3) + 165.87 * Math.Pow(DATA.Count(), -2.5) - 143.77 * Math.Pow(DATA.Count(), -2) +
                        67.727 * Math.Pow(DATA.Count(), -1.5) - 17.755 * Math.Pow(DATA.Count(), -1) + 4.323 * Math.Pow(DATA.Count(), -0.5) + 0.711;
                }

                else                                                            // Уровень значимости = 0.1 
                {
                    tableIrvinValue = -132.78 * Math.Pow(DATA.Count(), -3) + 224.24 * Math.Pow(DATA.Count(), -2.5) - 165.27 * Math.Pow(DATA.Count(), -2) +
                        68.614 * Math.Pow(DATA.Count(), -1.5) - 16.109 * Math.Pow(DATA.Count(), -1) + 3.693 * Math.Pow(DATA.Count(), -0.5) + 0.549;
                }

                if (calculatedIrvinValue > tableIrvinValue) DATA.Remove(suspectValue);
                else break;
            }
        }
        private void btnGo_Click(object sender, EventArgs e)
            {
            if (alfa == 0)
            {
                MessageBox.Show("Выберите критерий значимости!");
                return;
            }
            // Получить данные.
            string[,] values = LoadCsv(textBox1.Text);
            int num_rows = values.GetUpperBound(0) + 1;
            int num_cols = values.GetUpperBound(1) + 1;

            for (int r = 1; r < num_rows; r++)
            {
                if (values[r, 1] != null)
                {
                    DATA.Add(Convert.ToDouble(values[r, 1]));
                    DATA1.Add(values[r, 0]);
                }
                else
                {
                    continue;
                }
            }

            for (int i = 0; i < DATA.Count; i++)
            {
                if (DATA[i] > 55 || DATA[i] < -95)
                {
                    DATA.RemoveAt(i);
                }
            }

            if (DATA.Count < 14)
            {
                MessageBox.Show("Слишком мало значений");
                return;
            }

            // Показывать данные, чтобы показать, что у нас это есть.

            // Создаем заголовки столбцов.
            // Для этого примера предположим, что первая строка
            // содержит имена столбцов.
            dataGridView1.Columns.Clear();
            for (int c = 0; c < num_cols; c++)
                dataGridView1.Columns.Add(values[0, c], values[0, c]);

            // Добавьте данные.
            for (int r = 1; r < num_rows; r++)
            {
                dataGridView1.Rows.Add();
                for (int c = 0; c < num_cols; c++)
                {
                    dataGridView1.Rows[r - 1].Cells[c].Value = values[r, c];
                    //textBox2.Text = values[r, 1];
                }
            }

            CheckOnMisses();

            double[] array = new double[DATA.Count];
            
            for (int r = 0; r < DATA.Count; r++)
            {
                double x = DATA[r];
                array[r] = x;
                if (array[r] > max)
                {
                    // найден больший элемент
                    max = array[r];
                }
            }
            for (int r = 0; r < DATA.Count; r++)
            {
                double x = DATA[r];
                array[r] = x;
                if (array[r] < min)
                {
                    // найден меньший элемент
                    min = array[r];
                }
            }

            k = (int)(1 + 3.322 * Math.Log10(DATA.Count));
            double[] intervalStart = new double[k];
            double[] intervalEnd = new double[k];
            h = ((max - min) / k);
            h1 = Math.Round(h);
            
            for (int i = 0; i < k; i++)
            {
                if (i == 0)
                {
                    intervalStart[i] = min;
                    intervalEnd[i] = intervalStart[i] + h1;
                }
                else if (i == (k - 1))
                {
                    intervalStart[i] = intervalEnd[i-1];
                    intervalEnd[i] = max;
                }
                else
                {
                    intervalStart[i] = intervalEnd[i - 1];
                    intervalEnd[i] = intervalStart[i] + h1;
                }
            }

            int[] numArr = new int[k];
            int[] countA = new int[k];
            int[] intervaly = new int[k];
            for (int i = 0; i < k; i++)
            {
                numArr[i] = 0;
                countA[i] = 0;
                if (i == (k - 1))
                {
                    for (int r = 0; r < DATA.Count; r++)
                    {
                        double x = DATA[r];
                        array[r] = x;
                        if (array[r] >= intervalStart[i] && array[r] <= intervalEnd[i])
                        {
                            numArr[i] += 1;
                        }
                    }

                    double[] arrayA = new double[numArr[i]];
                    for (int r = 0; r < DATA.Count; r++)
                    {
                        double x = DATA[r];
                        array[r] = x;
                        if (array[r] >= intervalStart[i] && array[r] <= intervalEnd[i])
                        {
                            arrayA[countA[i]] = array[r];
                            countA[i] += 1;
                        }
                    }
                }
                else
                {
                    for (int r = 0; r < DATA.Count; r++)
                    {
                        double x = DATA[r];
                        array[r] = x;
                        if (array[r] >= intervalStart[i] && array[r] < intervalEnd[i])
                        {
                            numArr[i] += 1;
                        }
                    }

                    double[] arrayA = new double[numArr[i]];
                    for (int r = 0; r < DATA.Count; r++)
                    {
                        double x = DATA[r];
                        array[r] = x;
                        if (array[r] >= intervalStart[i] && array[r] < intervalEnd[i])
                        {
                            arrayA[countA[i]] = array[r];
                            countA[i] += 1;
                        }
                    }
                }
                intervaly[i] = numArr[i];
            }
            
            for (int r = 0; r < k; r++)
            {
                tmp.Add(intervaly[r]);
                tmp1.Add(intervalStart[r]);
                tmp2.Add(intervalEnd[r]);
            }

        

                int minValue;
                int minIndex;

            for (int i = 0; i < k; i++)
            {
                bool flag = false;
                minValue = tmp.Min();
                minIndex = tmp.IndexOf(minValue);
                if (minValue < 5)
                {
                    if (minIndex == 0)
                    {
                        tmp[1] += minValue;
                        tmp.RemoveAt(0);
                        tmp2[0] = tmp2[1];
                        tmp1.RemoveAt(1);
                        tmp2.RemoveAt(1);
                        flag = true;
                    }
                    if (minIndex == tmp.Count - 1 && flag == false)
                    {
                        tmp[tmp.Count - 2] += minValue;

                        tmp2[tmp.Count - 2] = tmp2[tmp.Count - 1];

                        tmp1.RemoveAt(tmp.Count - 1);
                        tmp2.RemoveAt(tmp.Count - 1);
                        tmp.RemoveAt(minIndex);
                        flag = true;
                    }
                    if (flag == false)
                    {
                        if (tmp[minIndex - 1] <= tmp[minIndex + 1])
                        {

                            tmp[minIndex - 1] += minValue;

                            tmp1[minIndex] = tmp1[minIndex - 1];
                            tmp.RemoveAt(minIndex);
                            tmp1.RemoveAt(minIndex - 1);
                            tmp2.RemoveAt(minIndex - 1);

                        }
                        else
                        {
                            tmp[minIndex + 1] += minValue;
                            tmp.RemoveAt(minIndex);
                            tmp2[minIndex] = tmp2[minIndex + 1];
                            tmp1.RemoveAt(minIndex + 1);
                            tmp2.RemoveAt(minIndex + 1);
                        }
                    }
                }
            }

            //мат.ожидание и тд
            double mCount = 0;
            for (int r = 0; r < DATA.Count; r++)
            {
                mCount = mCount + DATA[r];
            }
            M = mCount / (DATA.Count - 1);

            double dCount = 0;
            for (int r = 0; r < DATA.Count; r++)
            {
                dCount = dCount + Math.Pow(DATA[r] - M, 2);
            }
            D = dCount / (DATA.Count - 2);

            Sigma = Math.Sqrt(D);

            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();

            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();

            chart1.Series.Add(series1);
            chart1.Series.Add(series2); ;
            //chart1.Series[0].LegendText = "Эмпирические частоты";
            //chart1.Series[1].LegendText = "Теоретические частоты";
            chart1.Series[0].ChartType = SeriesChartType.Column;
            chart1.Series[1].ChartType = SeriesChartType.Spline;
            chart1.Series[0].BorderWidth = 5;
            chart1.Series[1].BorderWidth = 3;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();

            var val2 = new ChartValues<ObservablePoint>();

            for (int i = 0; i < tmp.Count; i++)
            {
                chart1.Series[0].Points.AddXY(i + 1, tmp[i]);
                val2.Add(new ObservablePoint(i + 1, tmp[i]));
            }
            SeriesCollection = new LiveCharts.SeriesCollection
                {

                    
                    new ColumnSeries
                    {
                        Title = "Эмперические частоты",
                        Values = val2
                    }
                };

            cartesianChart1.Series = SeriesCollection;
            btnGo.Enabled = false;
            button1.Visible = true;
            button1.Enabled = true;
            FileToolStripMenuItem.Enabled = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void FileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void открытьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Open();
            //btnGo.Visible = true;
            //btnGo.Enabled = true;
        }

        void Open()
        {
            string fileContent = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "C:\\Users\\KOLO BEE\\Desktop";
                openFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            if (fileContent != string.Empty)
            {
                //fileContent = fileContent.Replace('\n', '\r');
                string[] lines = fileContent.Split(new char[] { '\r' },
                    StringSplitOptions.RemoveEmptyEntries);

                if (lines[0].Contains(";"))
                {
                    textBox1.Text = fileContent;
                    btnGo.Visible = true;
                    btnGo.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Неправильный разделитель в первой строчке");
                }
                
            }
        }

        void SaveAs()
        {
            SaveFileDialog saveAsFileDialog = new SaveFileDialog();
            string fileContent = textBox1.Text;

            saveAsFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveAsFileDialog.FilterIndex = 1;
            saveAsFileDialog.RestoreDirectory = true;

            if (saveAsFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = saveAsFileDialog.FileName;
                File.WriteAllText(saveAsFileDialog.FileName, fileContent);
            }
        }
        void Save()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string fileContent = textBox1.Text;
            saveFileDialog.FileName = filePath;
            File.WriteAllText(saveFileDialog.FileName, fileContent);
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (File.Exists(filePath))
            {
                Save();
            }

            else
            {
                SaveAs();
            }
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAs();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {            

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        public LiveCharts.SeriesCollection SeriesCollection { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {       
            if (zakon == 0)
            {
                MessageBox.Show("Выберите закон распределения!");
                return;
            }
  
                SetSettings();

            if (zakon == 1)
            {
                double[] probability = new double[tmp.Count];
                int N = DATA.Count();
                arithmeticMean = DATA.Sum() / N;

                for (int i = 0; i < tmp.Count; i++)
                {
                    probability[i] = LaplasFunction(tmp1[i], tmp2[i]);
                    p.Add(probability[i]);
                    np.Add(probability[i] * N);
                    n.Add(tmp[i]);
                    nnp2.Add(Math.Pow(n[i] - np[i], 2));
                    nnp2np.Add(nnp2[i] / np[i]);
                }

                

                for (int i = 0;i < tmp.Count; i++)
                {
                    dataGridView2.Rows[i].Cells[1].Value = n[i];
                    dataGridView2.Rows[i].Cells[2].Value = p[i];
                    dataGridView2.Rows[i].Cells[3].Value = np[i];
                    dataGridView2.Rows[i].Cells[4].Value = nnp2[i].ToString("0.00");
                    dataGridView2.Rows[i].Cells[5].Value = nnp2np[i].ToString("0.00");
                }

                double findedX2 = SumColumnCell(5);                                 // Хи квадрат вычисленное (эксперементальное значение статистики)

                dataGridView2.Rows[tmp.Count].Cells[1].Value = SumColumnCell(1).ToString();
                dataGridView2.Rows[tmp.Count].Cells[2].Value = SumColumnCell(2).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[3].Value = SumColumnCell(3).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[4].Value = SumColumnCell(4).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[5].Value = findedX2.ToString("0.00");
                int countDegreesOfFreedom = DegreesOfFreedom(tmp.Count, 2);
                double tableX2 = PearsonsCriterion(countDegreesOfFreedom);
                richTextBox1.Text += "Минимальное значение: " + min;
                richTextBox1.Text += "\nМаксимальное значение: " + max;
                richTextBox1.Text += "\nКол-во интервалов: " + Convert.ToString(tmp.Count);
                richTextBox1.Text += "\nСреднее значение: " + arithmeticMean.ToString("0.00");
                richTextBox1.Text += "\nДисперсия: " + D.ToString("0.00");
                richTextBox1.Text += "\nСреднее квадратичное отклонение: " + Sigma.ToString("0.00");
                richTextBox1.Text += "\nКол-во степеней свободы: " + countDegreesOfFreedom;
                richTextBox1.Text += "\nX^2 вычисленное: " + findedX2;
                richTextBox1.Text += "\nX^2 табличное: " + tableX2;

                if (findedX2 <= tableX2)                                            // Вывод заключения 
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nбольше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо нормальному закону распределения\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nпринимается.";
                }
                else
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nменьше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо нормальному закону распределения\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nотвергается.";
                }  
                
            }
            if (zakon == 2)
            {

                double[] probability = new double[tmp.Count];
                int N = DATA.Count();
                arithmeticMean = DATA.Sum() / N;
                double h = 1 / (tmp2[tmp.Count - 1] - tmp1[0]); 
                for (int i = 0; i < tmp.Count; i++)
                {
                    probability[i] = h * (tmp2[i] - tmp1[0] - (tmp1[i] - tmp1[0]));
                    p.Add(probability[i]);
                    np.Add(probability[i] * N);
                    n.Add(tmp[i]);
                    nnp2.Add(Math.Pow(n[i] - np[i], 2));
                    nnp2np.Add(nnp2[i] / np[i]);
                }

                for (int i = 0; i < tmp.Count; i++)
                {
                    
                    dataGridView2.Rows[i].Cells[1].Value = n[i];
                    dataGridView2.Rows[i].Cells[2].Value = p[i];
                    dataGridView2.Rows[i].Cells[3].Value = np[i];
                    dataGridView2.Rows[i].Cells[4].Value = nnp2[i].ToString("0.00");
                    dataGridView2.Rows[i].Cells[5].Value = nnp2np[i].ToString("0.00");
                }

                double findedX2 = SumColumnCell(5);                                 // Хи квадрат вычисленное (эксперементальное значение статистики)

                dataGridView2.Rows[tmp.Count].Cells[1].Value = SumColumnCell(1).ToString();
                dataGridView2.Rows[tmp.Count].Cells[2].Value = SumColumnCell(2).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[3].Value = SumColumnCell(3).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[4].Value = SumColumnCell(4).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[5].Value = findedX2.ToString("0.00");
                int countDegreesOfFreedom = DegreesOfFreedom(tmp.Count, 2);
                double tableX2 = PearsonsCriterion(countDegreesOfFreedom);
                richTextBox1.Text += "Минимальное значение: " + min;
                richTextBox1.Text += "\nМаксимальное значение: " + max;
                richTextBox1.Text += "\nКол-во интервалов: " + Convert.ToString(tmp.Count);
                richTextBox1.Text += "\nСреднее значение: " + arithmeticMean.ToString("0.00");
                richTextBox1.Text += "\nДисперсия: " + D.ToString("0.00");
                richTextBox1.Text += "\nСреднее квадратичное отклонение: " + Sigma.ToString("0.00");
                richTextBox1.Text += "\nКол-во степеней свободы: " + countDegreesOfFreedom;
                richTextBox1.Text += "\nX^2 вычисленное: " + findedX2;
                richTextBox1.Text += "\nX^2 табличное: " + tableX2;

                if (findedX2 <= tableX2)                                            // Вывод заключения 
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nбольше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо равномерному закону распределения\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nпринимается.";
                }
                else
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nменьше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо равномерному закону распределения\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nотвергается.";
                }
            }
            if (zakon == 3)
            {
                int N = DATA.Count();

                double[] probability = new double[tmp.Count];
                arithmeticMean = DATA.Sum() / N;
                float[] fact = new float[tmp.Count];
                double Lyam;
                double[] avgIn = new double[tmp.Count];
                double sumAvg = 0;

                for (int i = 0; i < tmp.Count; i++)
                {
                    avgIn[i] = (int)(tmp1[i] + tmp2[i]) / 2;
                    sumAvg += avgIn[i] * tmp[i];
                }
                

                for (int i = 0; i < tmp.Count; i++)
                {
                    float BI = (float)Factorial(tmp[i]);
                    fact[i] = BI;
                    Lyam = tmp[i];

                    probability[i] = (double)(Math.Exp(-Lyam) * Math.Pow(Lyam, tmp[i])) / fact[i];
                    p.Add(probability[i]);
                    np.Add(probability[i] * N);
                    n.Add(tmp[i]);
                    nnp2.Add(Math.Pow(n[i] - np[i], 2));
                    nnp2np.Add(nnp2[i] / np[i]);
                }

                for (int i = 0; i < tmp.Count; i++)
                {
                    dataGridView2.Rows[i].Cells[1].Value = n[i];
                    dataGridView2.Rows[i].Cells[2].Value = p[i];
                    dataGridView2.Rows[i].Cells[3].Value = np[i];
                    dataGridView2.Rows[i].Cells[4].Value = nnp2[i].ToString("0.00");
                    dataGridView2.Rows[i].Cells[5].Value = nnp2np[i].ToString("0.00");
                }

                double findedX2 = SumColumnCell(5);                                 // Хи квадрат вычисленное (эксперементальное значение статистики)

                dataGridView2.Rows[tmp.Count].Cells[1].Value = SumColumnCell(1).ToString();
                dataGridView2.Rows[tmp.Count].Cells[2].Value = SumColumnCell(2).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[3].Value = SumColumnCell(3).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[4].Value = SumColumnCell(4).ToString("0.00");
                dataGridView2.Rows[tmp.Count].Cells[5].Value = findedX2.ToString("0.00");
                int countDegreesOfFreedom = DegreesOfFreedom(tmp.Count, 1);
                double tableX2 = PearsonsCriterion(countDegreesOfFreedom);
                richTextBox1.Text += "Минимальное значение: " + min;
                richTextBox1.Text += "\nМаксимальное значение: " + max;
                richTextBox1.Text += "\nКол-во интервалов: " + Convert.ToString(tmp.Count);
                richTextBox1.Text += "\nСреднее значение: " + arithmeticMean.ToString("0.00");
                richTextBox1.Text += "\nДисперсия: " + D.ToString("0.00");
                richTextBox1.Text += "\nСреднее квадратичное отклонение: " + Sigma.ToString("0.00");
                richTextBox1.Text += "\nКол-во степеней свободы: " + countDegreesOfFreedom;
                richTextBox1.Text += "\nX^2 вычисленное: " + findedX2;
                richTextBox1.Text += "\nX^2 табличное: " + tableX2;

                if (findedX2 <= tableX2)                                            // Вывод заключения 
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nбольше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо распределению Пуассона\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nпринимается.";
                }
                else
                {
                    label3.Text = "Табличное значение статистики (" + tableX2 + ") \nменьше вычисленного значения (" + findedX2 + "), \nто делаем вывод, что гипотеза H0\nпо распределению Пуассона\nс параметрами (" +
                         arithmeticMean.ToString("0.00") + "; " + D.ToString("0.00") + ") \nотвергается.";
                }
            }
            button1.Enabled = false;
            button2.Visible = true;
            button2.Enabled = true;
        }

        private void SetSettings()                                              // Функция настройки параметров таблицы и графика 

        {

            chart1.Series[0].LegendText = "Эмпирические частоты";
            chart1.Series[1].LegendText = "Теоретические частоты";

            dataGridView2.ColumnCount = 6;
            dataGridView2.RowCount = tmp.Count+1;

            for (int i = 0; i < tmp.Count; i++)
            {
                dataGridView2.Rows[i].Cells[0].Value = tmp1[i] + " - " + tmp2[i];
            }

            dataGridView2.Columns[0].HeaderText = "Интервал";
            dataGridView2.Columns[1].HeaderText = "Эмпирические частоты (ni)";
            dataGridView2.Columns[2].HeaderText = "Вероятность (pi)";
            dataGridView2.Columns[3].HeaderText = "Теоретические частоты (npi)";
            dataGridView2.Columns[4].HeaderText = "(ni - npi)^2";
            dataGridView2.Columns[5].HeaderText = "(ni - npi)^2/npi";

            dataGridView2.Rows[tmp.Count].Cells[0].Value = "Итого:";

        }

        private double PearsonsCriterion(int countDegreesOfFreedom)             // Функция для нахождения табличного значения критерия Пирсона 
        {
            double tableX2 = 0;

            for (int i = 4; i < X2_TABLE.Count; i += 4)                         // По известным уровнем значимости и кол-ву степеней свободу находи нужное табличное значение X^2 (критическое значение статистики) 
            {
                if (X2_TABLE[i] == countDegreesOfFreedom)
                {
                    if (alfa == 0.01)                                 // Уровень значимости = 0.01 
                    {
                        tableX2 = X2_TABLE[i + 3];
                    }
                    else if (alfa == 0.05)                            // Уровень значимости = 0.05 
                    {
                        tableX2 = X2_TABLE[i + 2];
                    }
                    else                                                        // Уровень значимости = 0.1 
                    {
                        tableX2 = X2_TABLE[i + 1];
                    }
                }
            }
            return tableX2;
        }
        private int DegreesOfFreedom(int m, int r)                              // Функция нахождения числа степеней свободы 
        {
            return m - r - 1;
        }

        private static BigInteger Factorial(double value)                                 // Функция вычисления факториала 
        {
            BigInteger result = new BigInteger(1);

            if (value == 1 || value == 0) return result;

            for (int i = 1; i < value; i++)
            {
                result += result * i;
            }
            return result;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var val1 = new ChartValues<ObservablePoint>();
            
            for (int i = 0; i < tmp.Count; i++)
            {
                val1.Add(new ObservablePoint(i + 1, np[i]));
                chart1.Series[1].Points.AddXY(i + 1, np[i]);
            }
            SeriesCollection.Add(new LineSeries { Title = "Теоретические частоты",Values = val1 });

            cartesianChart1.Series = SeriesCollection;

            button2.Enabled = false;
            button3.Visible = true;
            button3.Enabled = true;
        }

        private void elementHost1_ChildChanged(object sender, System.Windows.Forms.Integration.ChildChangedEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            {
                SaveFileDialog saveFD = new SaveFileDialog();

                bool fileError = false;                                             // Флаг на ошибки 

                if (radioButton1.Checked == true)                                   // В зависимости от выбранного формата сохранения устанавливаем соответствующий фильтр на форматы и стандартное имя файла 
                {
                    saveFD.Filter = "DOCX (*.docx)|*.docx";
                    saveFD.FileName = "Result.docx";
                }

                if (radioButton2.Checked == true)
                {
                    saveFD.Filter = "PDF (*.pdf)|*.pdf";
                    saveFD.FileName = "Result.pdf";
                }

                if (saveFD.ShowDialog() == DialogResult.OK)                         // Если выбран файл и нажата кнопка ОК 
                {
                    if (File.Exists(saveFD.FileName))                               // Если файл введенным именем существует 
                    {
                        try
                        {
                            File.Delete(saveFD.FileName);                           // Стираем его и создаем новый 
                        }
                        catch (IOException ex)                                      // Если постирать не удалось - кидаемся ошибками 
                        {
                            fileError = true;                                       // И устанавливаем флаг 
                            MessageBox.Show("Невозможно сохранить данные. Ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    if (!fileError)                                                 // Если все чудненько и ошибок не возникло 
                    {
                        if (radioButton1.Checked == true || radioButton2.Checked == true)                                                                        // Если выбрано сохранение в docx или pdf 
                        {
                            var application = new Word.Application();               // Создаем экземпляр приложения ворда 
                            application.Visible = false;                            // Он отображается в фоновом режиме 

                            Word.Document newDocument = application.Documents.Add();// Создаем новый документ 

                            Word.Paragraph title = newDocument.Paragraphs.Add();    // Создаем новый параграф-заголовок 

                            Word.Range titleRange = title.Range;                    // Получаем его range (отвечает за непосредственный доступ к тексту) и настраиваем с его помощью текст 

                            titleRange.Text = "Результат проверки гипотезы о распределении случайной величины по " + label7.Text + " с критерием значимости " + label5.Text;

                            titleRange.Font.Size = 15;

                            titleRange.Font.Bold = 1;

                            title.SpaceAfter = 20;

                            titleRange.InsertParagraphAfter();


                            Word.Paragraph subTitleTable = newDocument.Paragraphs.Add();                                                                        // Делаем тоже самое для подзаголовка 

                            Word.Range subTitleTableRange = subTitleTable.Range;

                            subTitleTableRange.Text = "Таблица расчетов";

                            subTitleTableRange.Font.Size = 12;

                            subTitleTableRange.Font.Bold = 1;

                            subTitleTable.SpaceAfter = 10;

                            subTitleTableRange.InsertParagraphAfter();


                            Word.Paragraph tableParagraph = newDocument.Paragraphs.Add();                                                                       // Параграф для таблицы 

                            Word.Range tableRange = tableParagraph.Range;

                            Word.Table table = newDocument.Tables.Add(tableRange, dataGridView2.RowCount + 1, dataGridView2.ColumnCount);

                            table.Borders.InsideLineStyle = table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            tableRange.InsertParagraphAfter();


                            Word.Range cellRange;                                   // Range для клеток таблицы

                            for (int i = 0; i < dataGridView2.ColumnCount; i++)      // Заполняем шапку таблицы 
                            {
                                cellRange = table.Cell(1, i + 1).Range;
                                cellRange.Text = dataGridView2.Columns[i].HeaderText;
                            }

                            table.Rows[1].Range.Font.Bold = 1;                      // Делаем в ней текст жирным и фон серым 

                            table.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

                            for (int i = 0; i < dataGridView2.RowCount; i++)         // Заполняем тело таблицы 
                            {
                                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                                {
                                    cellRange = table.Cell(i + 2, j + 1).Range;
                                    cellRange.Text = Convert.ToString(dataGridView2[j, i].Value);
                                }
                            }


                            Word.Paragraph subTitleChart = newDocument.Paragraphs.Add();                                                                        // Создаем параграфы с остальной информацией и параграфы-заголовки для них 

                            Word.Range subTitleChartRange = subTitleChart.Range;

                            subTitleChartRange.Text = "График";

                            subTitleChartRange.Font.Size = 12;

                            subTitleChartRange.Font.Bold = 1;

                            subTitleChartRange.InsertParagraphAfter();

                            chart1.SaveImage("Chart.png", System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);                                  // Сохраняем изображение графика 


                            Word.Paragraph сhart = newDocument.Paragraphs.Add();

                            Word.Range сhartRange = сhart.Range;

                            сhartRange.InlineShapes.AddPicture(Environment.CurrentDirectory + @"\Chart.png");                                                   // Вставляем его в соответствующий параграф 

                            File.Delete("Chart.png");                               // Удаляем изображение чтобы не занимало места 


                            Word.Paragraph subTitleCalculatedValues = newDocument.Paragraphs.Add();

                            Word.Range subTitleCalculatedValuesRange = subTitleCalculatedValues.Range;

                            subTitleCalculatedValuesRange.Text = "Вычисленные данные:";

                            subTitleCalculatedValuesRange.Font.Size = 12;

                            subTitleCalculatedValuesRange.Font.Bold = 1;

                            subTitleCalculatedValuesRange.InsertParagraphAfter();


                            Word.Paragraph сalculatedValues = newDocument.Paragraphs.Add();

                            Word.Range сalculatedValuesRange = сalculatedValues.Range;

                            сalculatedValuesRange.Text = richTextBox1.Text;

                            сalculatedValuesRange.Font.Size = 12;

                            сalculatedValuesRange.InsertParagraphAfter();


                            Word.Paragraph subTitleConclusion = newDocument.Paragraphs.Add();

                            Word.Range subTitleConclusionRange = subTitleConclusion.Range;

                            subTitleConclusionRange.Text = "Вывод:";

                            subTitleConclusionRange.Font.Size = 12;

                            subTitleConclusionRange.Font.Bold = 1;

                            subTitleConclusionRange.InsertParagraphAfter();


                            Word.Paragraph conclusion = newDocument.Paragraphs.Add();

                            Word.Range conclusionRange = conclusion.Range;

                            conclusionRange.Text = label3.Text;

                            conclusionRange.Font.Size = 12;

                            conclusionRange.InsertParagraphAfter();


                            if (radioButton1.Checked == true) newDocument.SaveAs2(saveFD.FileName);                                                             // Сохраняем в выбранном формате 
                            else newDocument.SaveAs2(saveFD.FileName, Word.WdSaveFormat.wdFormatPDF);


                            newDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);// Закрываем документ без сохранения (оно было выше)                           

                            newDocument = null;

                            application.Quit();                                     // Закрывем приложение 

                            application = null;

                        }

                    }

                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tmp.Clear();
            k = 0;
            min = int.MaxValue;
            max = int.MinValue;
            h1 = 0;
            h = 0;
            
            tmp1.Clear();
            tmp2.Clear();
            DATA.Clear();
            DATA1.Clear();

            p.Clear();
            np.Clear();
            n.Clear();
            nnp2.Clear();
            nnp2np.Clear();

            dataGridView1.Rows.Clear();                                         // Очитстка таблиц 
            dataGridView1.Columns.Clear();
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            textBox1.Text = "";
            richTextBox1.Text = "";

            label3.Text = "Вывод";
            

            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton7.Checked = false;
            radioButton8.Checked = false;

            label5.Text = "alfa";
            label7.Text = "Выбранный закон";

            zakon = 0;
            alfa = 0;

            chart1.Series.Clear();
            cartesianChart1.Series.Clear();

            btnGo.Enabled = false;
            button1.Enabled= false;
            button2.Enabled = false;
            button3.Enabled = false;
            FileToolStripMenuItem.Enabled = true;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {
            
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
            
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            alfa = 0.05;
            label5.Text = "" + alfa;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            alfa = 0.1;
            label5.Text = "" + alfa;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            alfa = 0.01;
            label5.Text = "" + alfa;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            zakon = 1;
            label7.Text = "Нормальный закон";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            zakon = 2;
            label7.Text = "Равномерный закон";
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            zakon = 3;
            label7.Text = "По Пуассону";
        }

        public double SumColumnCell(int columnNumber)                          // Функция подсчета суммы всех значений в столбце 
        {
            double sum = 0;

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                sum += Convert.ToDouble(dataGridView2[columnNumber, i].Value);
            }
            return sum;

        }

        private double LaplasFunction(double intervalStart, double intervalEnd) // Функция Лапласа для вычисления вероятности попадания случайной величины в интервал 
        {

            double laplasFuncValue1 = Math.Round((intervalEnd - arithmeticMean) / Sigma, 2);                                                  // Вычисляем первую дробь из формулы Лапласа 
            double laplasFuncValue2 = Math.Round((intervalStart - arithmeticMean) / Sigma, 2);                                                // Вычисляем вторую дробь из формулы Лапласа 
            double laplasTableValue1 = 0;                                       // Переменные для записи значений из таблицы, которые будут соответствовать вычисленным выше 
            double laplasTableValue2 = 0;
            bool firstValueFinded = false;                                      // Флаги для контроля, что в таблице нашлось значение для вычисленных выше значений 
            bool secondValueFinded = false;

            for (int i = 0; i < LAPLAS_TABLE.Count(); i += 2)                    // Подбираем значения из таблицы для вычисленных значений 
            {
                if (LAPLAS_TABLE[i] == Math.Abs(laplasFuncValue1))
                {
                    laplasTableValue1 = LAPLAS_TABLE[i + 1];
                    firstValueFinded = true;
                }

                if (LAPLAS_TABLE[i] == Math.Abs(laplasFuncValue2))
                {
                    laplasTableValue2 = LAPLAS_TABLE[i + 1];
                    secondValueFinded = true;
                }
            }

            if (firstValueFinded == false || secondValueFinded == false)         // Если для какого-либо вычисленного значения в таблице не нашлось соответствия, то выбираем ближайшее для которого соответствие есть 
            {

                if (firstValueFinded == false)
                {
                    double closestValue = 0;
                    double temp = LAPLAS_TABLE.Max();

                    for (int i = 0; i < LAPLAS_TABLE.Count(); i += 2)
                    {
                        if (Math.Abs(LAPLAS_TABLE[i] - Math.Abs(laplasFuncValue1)) < temp)
                        {
                            closestValue = LAPLAS_TABLE[i + 1];
                            temp = Math.Abs(LAPLAS_TABLE[i] - Math.Abs(laplasFuncValue1));
                        }
                    }

                    laplasTableValue1 = closestValue;

                }

                if (secondValueFinded == false)
                {
                    double closestValue = 0;
                    double temp = LAPLAS_TABLE.Max();

                    for (int i = 0; i < LAPLAS_TABLE.Count(); i += 2)
                    {
                        if (Math.Abs(LAPLAS_TABLE[i] - Math.Abs(laplasFuncValue2)) < temp)
                        {
                            closestValue = LAPLAS_TABLE[i + 1];
                            temp = Math.Abs(LAPLAS_TABLE[i] - Math.Abs(laplasFuncValue2));
                        }
                    }

                    laplasTableValue2 = closestValue;

                }

            }

            if (laplasFuncValue1 > 0 && laplasFuncValue2 > 0) return laplasTableValue1 - laplasTableValue2;                                                 // т.к. вид возвращаемого выражения зависит от знаков в вычисленных дробях 
            else if (laplasFuncValue1 < 0 && laplasFuncValue2 > 0) return -laplasTableValue1 - laplasTableValue2;
            else if (laplasFuncValue1 > 0 && laplasFuncValue2 < 0) return laplasTableValue1 + laplasTableValue2;
            return -laplasTableValue1 + laplasTableValue2;
        }
    }

}
