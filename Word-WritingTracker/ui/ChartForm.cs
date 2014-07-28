using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Word_WritingTracker
{
    public partial class ChartForm : Form
    {
        public ChartForm()
        {
            InitializeComponent();
        }

        private void setDefaults()
        {
            comboBoxTimeSpan.SelectedIndex = 0;
            comboBoxChartType.SelectedIndex = 0;
            dateTimePickerEnd.Value = DateTime.Now;
            dateTimePickerStart.Value = dateTimePickerEnd.Value.StartOfMonth();

        }

        private void setWordsPerDayChart()
        {
            Tuple<DateTime, DateTime> range = GetTimeSpanFromComboBox();
            this.chart.ChartAreas[0].AxisX.LabelStyle.Interval = (range.Item2 - range.Item1).Days / 5.0;
            this.chart.ChartAreas[0].AxisX.LabelStyle.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Days;


            this.chart.Series.Clear();


            this.chart.Legends[0].Enabled = true;
            this.chart.Titles[0].Text = "Words / Day";
            this.chart.ChartAreas[0].AxisX.Title = "Date";

            Dictionary<String, List<Tuple<DateTime, int>>> dict = Util.GetDailyWordCount(range.Item1, range.Item2);
            foreach (String projectName in dict.Keys)
            {
                List<Tuple<DateTime, int>> dataList = new List<Tuple<DateTime, int>>();
                if (!dict.TryGetValue(projectName, out dataList))
                    System.Diagnostics.Debug.WriteLine("Failed to get data list value");

                var series = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = projectName,
                    ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedColumn,
                    XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Date
                };

                this.chart.Series.Add(series);

                foreach (Tuple<DateTime, int> tuple in dataList)
                {
                    // add a datapoint with a tool tip
                    System.Windows.Forms.DataVisualization.Charting.DataPoint point = new System.Windows.Forms.DataVisualization.Charting.DataPoint();
                    point.SetValueXY(tuple.Item1, tuple.Item2);
                    point.ToolTip = String.Format("{0:MM/dd/yyyy} - {1}", tuple.Item1, tuple.Item2);
                    series.Points.Add(point);
                }
            }
            this.chart.Invalidate();
        }

        private void setWordsPerProjectChart()
        {
            // clear data
            this.chart.Series.Clear();

            // set interval
            this.chart.ChartAreas[0].AxisX.LabelStyle.Interval = 1;
            this.chart.ChartAreas[0].AxisX.LabelStyle.IntervalType = System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType.Auto;

            // set new series
            var series = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Total Words",
                ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column,
                XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String
            };

            this.chart.Series.Add(series);

            this.chart.Legends[0].Enabled = false;
            this.chart.Titles[0].Text = "Words / Project";
            this.chart.ChartAreas[0].AxisX.Title = "Project Name";

            var dict = Util.GetLastMetricOfDayForTrackedProjects();

            foreach (TrackedFile tf in dict.Keys)
            {
                var metricList = new List<Metric>();
                if (!dict.TryGetValue(tf, out metricList))
                    System.Diagnostics.Debug.WriteLine("Failed to get metricList from dictionary");

                int wordCount = metricList.OrderByDescending(m => m.TimeStamp).First().WordCount;
                // set up data point with a tool tip
                System.Windows.Forms.DataVisualization.Charting.DataPoint point = new System.Windows.Forms.DataVisualization.Charting.DataPoint();
                point.SetValueXY(tf.ProjectName, wordCount);
                point.ToolTip = String.Format("{0} - {1}", tf.ProjectName, wordCount);
                series.Points.Add(point);

            }
            this.chart.Invalidate();
        }

        private Tuple<DateTime, DateTime> GetTimeSpanFromComboBox()
        {
            DateTime start, end;
            end = DateTime.Now;
            switch ((string)comboBoxTimeSpan.SelectedItem)
            {
                case "This Week":
                    start = end.StartOfWeek(DayOfWeek.Monday);
                    break;
                case "This Month":
                    start = end.StartOfMonth();
                    break;
                case "This Quarter":
                    start = end.StartOfQuarter();
                    break;
                case "This Year":
                    start = end.StartOfYear();
                    break;
                case "Last 7 Days":
                    start = end.AddDays(-7);
                    break;
                case "Last 14 Days":
                    start = end.AddDays(-14);
                    break;
                case "Last 21 Days":
                    start = end.AddDays(-21);
                    break;
                case "Last 28 Days":
                    start = end.AddDays(-28);
                    break;
                case "Last 365 Days":
                    start = end.AddDays(-365);
                    break;
                case "Custom":
                    start = dateTimePickerStart.Value;
                    end = dateTimePickerEnd.Value;
                    break;
                default:
                    // default to week
                    start = end.StartOfWeek(DayOfWeek.Monday);
                    break;
            }
            return new Tuple<DateTime, DateTime>(start, end);
        }

        private void ChartForm_Load(object sender, EventArgs e)
        {
            setDefaults();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Image Files (*.bmp;*.gif;*.jpg;*.jpeg;*.png;*.tif;*.tiff)|*.bmp;*.gif;*.jpg;*.jpeg;*.png;*.tif;*.tiff",
                FileName = comboBoxTimeSpan.SelectedItem + " Chart.png"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string path = sfd.FileName;
                System.Windows.Forms.DataVisualization.Charting.ChartImageFormat format;
                switch (System.IO.Path.GetExtension(path))
                {
                    case ".bmp":
                        format = System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Bmp;
                        break;
                    case ".gif":
                        format = System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Gif;
                        break;
                    case ".jpg":
                    case ".jpeg":
                        format = System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg;
                        break;
                    case ".png":
                        format = System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png;
                        break;
                    case ".tif":
                    case ".tiff":
                        format = System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Tiff;
                        break;
                    default:
                        throw new NotImplementedException("This extension hasn't been implemented for a save");
                }

                this.chart.SaveImage(path, format);
            }
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBoxChartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            switch ((string)cb.SelectedItem)
            {
                //case "Words / Day":
                //    break;  
                case "Words / Project":
                    comboBoxTimeSpan.Enabled = false;
                    labelTimeSpan.Enabled = false;
                    groupBoxCustomTime.Enabled = false;
                    setWordsPerProjectChart();
                    break;
                default:
                    // default to words / day
                    comboBoxTimeSpan.Enabled = true;
                    labelTimeSpan.Enabled = true;
                    comboBoxTimeSpan_SelectedIndexChanged(comboBoxTimeSpan, new EventArgs());
                    break;
            }
        }

        private void comboBoxTimeSpan_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            switch ((string)cb.SelectedItem)
            {
                case "Custom":
                    groupBoxCustomTime.Enabled = true;
                    break;
                default:
                    groupBoxCustomTime.Enabled = false;
                    break;
            }

            setWordsPerDayChart();
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            comboBoxTimeSpan_SelectedIndexChanged(comboBoxTimeSpan, new EventArgs());
        }

        private void dateTimePickerEnd_ValueChanged(object sender, EventArgs e)
        {
            comboBoxTimeSpan_SelectedIndexChanged(comboBoxTimeSpan, new EventArgs());
        }
    }
}
