using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace CSVChat
{
    public partial class FrmMain : Form
    {
        private DataTable csvData = new DataTable();
        private ContextMenuStrip chartContextMenu;
        private Series selectedSeries;
        private ToolTip dataPointToolTip = new ToolTip();
        private bool isMouseInChart = false;
        private DateTime lastMouseMoveTime = DateTime.MinValue;

        // 十字线注解
        private LineAnnotation verticalLine;
        private LineAnnotation horizontalLine;

        private BackgroundWorker loadingWorker;
        private int totalLines = 0;
        private int processedLines = 0;

        public FrmMain()
        {
            InitializeComponent();
            InitializeChartContextMenu();
            InitializeCrosshair();
            chart_Main.Series.Clear();
            InitializeBackgroundWorker();

            // 启用拖放功能
            this.AllowDrop = true;
            this.DragEnter += FrmMain_DragEnter;
            this.DragDrop += FrmMain_DragDrop;
        }

        #region 基础功能
        private void btn_OpenCSVFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|Logfile files (*.logfile)|*.logfile|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtb_CSVFilePath.Text = openFileDialog.FileName;
                    LoadCSVHeaders(openFileDialog.FileName);
                }
            }
        }

        private void LoadCSVHeaders(string filePath)
        {
            try
            {
                clb_Name.Items.Clear();
                csvData = new DataTable();

                string[] headers = File.ReadLines(filePath).First().Split(',');

                foreach (string header in headers)
                {
                    clb_Name.Items.Add(header.Trim());
                }

                // 自动勾选前两列（可选）
                //if (clb_Name.Items.Count >= 2)
                //{
                //    clb_Name.SetItemChecked(0, true);
                //    clb_Name.SetItemChecked(1, true);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载CSV文件头失败: {ex.Message}");
            }
        }

        private void btn_Load_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtb_CSVFilePath.Text) || !File.Exists(txtb_CSVFilePath.Text))
            {
                MessageBox.Show("请先选择有效的CSV文件");
                return;
            }

            if (clb_Name.CheckedItems.Count == 0)
            {
                MessageBox.Show("请至少选择一列数据");
                return;
            }

            LoadCSVDataAndPlot();
        }

        private void LoadCSVDataAndPlot()
        {
            try
            {
                SetControlsEnabled(false);
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                loadingWorker.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载数据失败: {ex.Message}");
                SetControlsEnabled(true);
                progressBar1.Visible = false;
            }
        }

        private void LoadingWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string filePath = txtb_CSVFilePath.Text;
            string[] csvLines = File.ReadAllLines(filePath);
            totalLines = csvLines.Length - 1;

            // 在UI线程初始化图表
            chart_Main.Invoke((MethodInvoker)delegate {
                chart_Main.SuspendLayout();
                chart_Main.Series.Clear();
                chart_Main.ChartAreas[0].AxisX.Title = "行号";
                chart_Main.ChartAreas[0].AxisY.Title = "值";

                // 启用缩放和拖拽
                chart_Main.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart_Main.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
                chart_Main.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                chart_Main.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            });

            // 预创建系列
            var seriesDict = new Dictionary<string, Series>();
            for (int j = 0; j < clb_Name.Items.Count; j++)
            {
                if (clb_Name.GetItemChecked(j))
                {
                    string columnName = clb_Name.Items[j].ToString();
                    var series = new Series(columnName)
                    {
                        ChartType = SeriesChartType.Line,
                        XValueType = ChartValueType.Double,
                        YValueType = ChartValueType.Double,
                        ShadowOffset = 0,
                        BorderWidth = 1
                    };
                    seriesDict[columnName] = series;

                    chart_Main.Invoke((MethodInvoker)delegate {
                        chart_Main.Series.Add(series);
                    });
                }
            }

            // 批量添加数据点
            int maxPoints = 5000;
            int step = Math.Max(1, totalLines / maxPoints);
            processedLines = 0;

            // 收集所有要添加的点
            var pointsToAdd = new Dictionary<string, List<DataPoint>>();
            foreach (var kvp in seriesDict)
            {
                pointsToAdd[kvp.Key] = new List<DataPoint>();
            }

            for (int i = 1; i < csvLines.Length; i += step)
            {
                if (loadingWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                if (i >= csvLines.Length) break;

                string[] fields = csvLines[i].Split(',');

                foreach (var kvp in seriesDict)
                {
                    int colIndex = clb_Name.Items.IndexOf(kvp.Key);
                    if (fields.Length > colIndex && double.TryParse(fields[colIndex], out double value))
                    {
                        pointsToAdd[kvp.Key].Add(new DataPoint(i, value));
                    }
                }

                processedLines += step;
                int progress = (int)((double)processedLines / totalLines * 100);
                loadingWorker.ReportProgress(progress);
            }

            // 一次性添加所有点（使用循环替代AddRange）
            chart_Main.Invoke((MethodInvoker)delegate {
                foreach (var kvp in seriesDict)
                {
                    foreach (var point in pointsToAdd[kvp.Key])
                    {
                        kvp.Value.Points.Add(point);
                    }
                }
            });
        }

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            if (loadingWorker.IsBusy)
            {
                loadingWorker.CancelAsync();
            }

            for (int i = 0; i < clb_Name.Items.Count; i++)
            {
                clb_Name.SetItemChecked(i, false);
            }

            chart_Main.Series.Clear();
            lbl_Status.Text = "就绪";
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            chart_Main.ChartAreas[0].AxisX.Title = "行号";
            chart_Main.ChartAreas[0].AxisY.Title = "值";
            chart_Main.Titles.Clear();
        }
        #endregion

        #region 右键菜单功能
        private void InitializeChartContextMenu()
        {
            chartContextMenu = new ContextMenuStrip();

            ToolStripMenuItem useSecondaryAxisItem = new ToolStripMenuItem("使用次坐标轴");
            useSecondaryAxisItem.Click += UseSecondaryAxisItem_Click;

            ToolStripMenuItem resetAxisItem = new ToolStripMenuItem("重置为主坐标轴");
            resetAxisItem.Click += ResetAxisItem_Click;

            chartContextMenu.Items.Add(useSecondaryAxisItem);
            chartContextMenu.Items.Add(resetAxisItem);

            chart_Main.ContextMenuStrip = chartContextMenu;
            chart_Main.MouseDown += Chart_Main_MouseDown;
        }

        private void Chart_Main_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                HitTestResult result = chart_Main.HitTest(e.X, e.Y);
                if (result.ChartElementType == ChartElementType.LegendItem ||
                    result.ChartElementType == ChartElementType.DataPoint)
                {
                    selectedSeries = result.Series;
                    chartContextMenu.Enabled = true;
                }
                else
                {
                    selectedSeries = null;
                    chartContextMenu.Enabled = false;
                }
            }
        }

        private void UseSecondaryAxisItem_Click(object sender, EventArgs e)
        {
            if (selectedSeries != null)
            {
                if (chart_Main.ChartAreas[0].AxisY2 == null)
                {
                    chart_Main.ChartAreas[0].AxisY2 = new Axis();
                    chart_Main.ChartAreas[0].AxisY2.Title = "次坐标轴";
                    chart_Main.ChartAreas[0].AxisY2.LineColor = Color.Red;
                    chart_Main.ChartAreas[0].AxisY2.MajorGrid.LineColor = Color.LightPink;
                    chart_Main.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.Red;
                }

                selectedSeries.YAxisType = AxisType.Secondary;
                chart_Main.ChartAreas[0].AxisY2.Enabled = AxisEnabled.True;
                chart_Main.ChartAreas[0].AxisY2.LabelAutoFitMinFontSize = 8;
                chart_Main.ChartAreas[0].AxisY2.IsStartedFromZero = false;
                chart_Main.ChartAreas[0].AxisY.IsStartedFromZero = false;
                chart_Main.ChartAreas[0].RecalculateAxesScale();
                chart_Main.Invalidate();
            }
        }

        private void ResetAxisItem_Click(object sender, EventArgs e)
        {
            if (selectedSeries != null)
            {
                selectedSeries.YAxisType = AxisType.Primary;
                chart_Main.Invalidate();
            }
        }
        #endregion

        #region 十字线功能
        private void InitializeCrosshair()
        {
            chart_Main.Annotations.Clear();

            verticalLine = new LineAnnotation();
            verticalLine.AxisX = chart_Main.ChartAreas[0].AxisX;
            verticalLine.AxisY = chart_Main.ChartAreas[0].AxisY;
            verticalLine.LineColor = Color.Red;
            verticalLine.LineWidth = 2;
            verticalLine.LineDashStyle = ChartDashStyle.Solid;
            verticalLine.IsInfinitive = true;
            verticalLine.ClipToChartArea = "ChartArea1";
            verticalLine.Visible = false;

            horizontalLine = new LineAnnotation();
            horizontalLine.AxisX = chart_Main.ChartAreas[0].AxisX;
            horizontalLine.AxisY = chart_Main.ChartAreas[0].AxisY;
            horizontalLine.LineColor = Color.Red;
            horizontalLine.LineWidth = 2;
            horizontalLine.LineDashStyle = ChartDashStyle.Solid;
            horizontalLine.IsInfinitive = true;
            horizontalLine.ClipToChartArea = "ChartArea1";
            horizontalLine.Visible = false;

            chart_Main.Annotations.Add(verticalLine);
            chart_Main.Annotations.Add(horizontalLine);

            chart_Main.MouseMove += Chart_Main_MouseMove;
            chart_Main.MouseEnter += Chart_Main_MouseEnter;
            chart_Main.MouseLeave += Chart_Main_MouseLeave;
            chart_Main.MouseClick += Chart_Main_MouseClick;

            chart_Main.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart_Main.ChartAreas[0].CursorY.IsUserEnabled = true;
        }

        private void ReinitializeCrosshair()
        {
            if (verticalLine != null && horizontalLine != null)
            {
                chart_Main.Annotations.Clear();
                chart_Main.Annotations.Add(verticalLine);
                chart_Main.Annotations.Add(horizontalLine);
                chart_Main.Invalidate();
            }
        }

        private void Chart_Main_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                var area = chart_Main.ChartAreas[0];
                double xValue = area.AxisX.PixelPositionToValue(e.X);
                double yValue = area.AxisY.PixelPositionToValue(e.Y);
                Console.WriteLine($"点击位置: X={xValue:F2}, Y={yValue:F2}");
            }
        }

        private void Chart_Main_MouseEnter(object sender, EventArgs e)
        {
            isMouseInChart = true;
            Console.WriteLine("鼠标进入图表区域");
        }

        private void Chart_Main_MouseLeave(object sender, EventArgs e)
        {
            isMouseInChart = false;
            verticalLine.Visible = false;
            horizontalLine.Visible = false;
            dataPointToolTip.Hide(chart_Main);
            chart_Main.Invalidate();
            Console.WriteLine("鼠标离开图表区域");
        }

        private void Chart_Main_MouseMove(object sender, MouseEventArgs e)
        {
            if ((DateTime.Now - lastMouseMoveTime).TotalMilliseconds < 20)
                return;
            lastMouseMoveTime = DateTime.Now;

            if (!isMouseInChart || chart_Main.Series.Count == 0)
                return;

            try
            {
                var area = chart_Main.ChartAreas[0];
                double xValue = area.AxisX.PixelPositionToValue(e.X);
                double yValue = area.AxisY.PixelPositionToValue(e.Y);

                verticalLine.X = xValue;
                horizontalLine.Y = yValue;
                verticalLine.Visible = true;
                horizontalLine.Visible = true;

                StringBuilder tooltipText = new StringBuilder();
                double xRangeThreshold = (area.AxisX.Maximum - area.AxisX.Minimum) / 30.0;

                foreach (Series series in chart_Main.Series)
                {
                    if (series.Points.Count == 0) continue;

                    DataPoint nearestPoint = null;
                    double minDistance = double.MaxValue;

                    foreach (DataPoint point in series.Points)
                    {
                        double distance = Math.Abs(point.XValue - xValue);
                        if (distance < minDistance && distance < xRangeThreshold)
                        {
                            minDistance = distance;
                            nearestPoint = point;
                        }
                    }

                    if (nearestPoint != null)
                    {
                        tooltipText.AppendLine($"{series.Name}: {nearestPoint.YValues[0]:F2}");
                    }
                }

                if (tooltipText.Length > 0)
                {
                    Point tooltipPosition = new Point(e.X + 20, e.Y + 20);
                    dataPointToolTip.Show(tooltipText.ToString().Trim(), chart_Main, tooltipPosition, 1000);
                }
                else
                {
                    dataPointToolTip.Hide(chart_Main);
                }

                chart_Main.Invalidate();
                Console.WriteLine($"鼠标移动: X={xValue:F2}, Y={yValue:F2}, 十字线可见={verticalLine.Visible}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MouseMove error: {ex.Message}");
            }
        }
        #endregion

        #region 测试按钮
        private void btn_TestCrosshair_Click(object sender, EventArgs e)
        {
            if (verticalLine != null)
            {
                verticalLine.X = 100;
                verticalLine.Visible = true;
                horizontalLine.Y = 50;
                horizontalLine.Visible = true;
                chart_Main.Invalidate();
                MessageBox.Show("十字线测试完成！如果看到红色十字线说明功能正常。");
            }
        }
        #endregion

        #region 后台工作线程
        private void InitializeBackgroundWorker()
        {
            loadingWorker = new BackgroundWorker();
            loadingWorker.WorkerReportsProgress = true;
            loadingWorker.WorkerSupportsCancellation = true;
            loadingWorker.DoWork += LoadingWorker_DoWork;
            loadingWorker.ProgressChanged += LoadingWorker_ProgressChanged;
            loadingWorker.RunWorkerCompleted += LoadingWorker_RunWorkerCompleted;
        }

        private void LoadingWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lbl_Status.Text = $"正在加载数据: {e.ProgressPercentage}%";
        }

        private void LoadingWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (e.Error != null)
                {
                    MessageBox.Show($"加载数据时出错: {e.Error.Message}");
                    progressBar1.Visible = false;
                }
                else if (e.Cancelled)
                {
                    lbl_Status.Text = "加载已取消";
                    progressBar1.Visible = false;
                }
                else
                {
                    lbl_Status.Text = "加载完成";
                    progressBar1.Value = 100;
                    progressBar1.Style = ProgressBarStyle.Continuous; // 确保显示为完成状态

                    chart_Main.Invoke((MethodInvoker)delegate {
                        chart_Main.ResumeLayout();
                        ReinitializeCrosshair();
                    });
                }
            }
            finally
            {
                SetControlsEnabled(true);
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            btn_OpenCSVFile.Enabled = enabled;
            btn_Load.Enabled = enabled;
            btn_Clear.Enabled = enabled;
            clb_Name.Enabled = enabled;
        }
        #endregion

        #region 保存图表
        private void btn_SavePlot_Click(object sender, EventArgs e)
        {
            try
            {
                // 设置保存对话框的默认文件名和过滤器
                saveFileDialog_Plot.FileName = $"ChartExport_{DateTime.Now:yyyyMMdd_HHmmss}";    // 默认文件名,使用精确到秒的时间戳作为文件名
                //saveFileDialog_Plot.Filter = "PNG 图片 (*.png)|*.png|JPEG 图片 (*.jpg;*.jpeg)|*.jpg;*.jpeg|BMP 图片 (*.bmp)|*.bmp|PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
                saveFileDialog_Plot.Filter = "PNG 图片 (*.png)|*.png|JPEG 图片 (*.jpg;*.jpeg)|*.jpg;*.jpeg|所有文件 (*.*)|*.*";
                saveFileDialog_Plot.FilterIndex = 1; // 默认选择PNG格式
                saveFileDialog_Plot.RestoreDirectory = true;

                if (saveFileDialog_Plot.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog_Plot.FileName;
                    string extension = Path.GetExtension(filePath).ToLower();

                    switch (extension)
                    {
                        case ".png":
                            chart_Main.SaveImage(filePath, ChartImageFormat.Png);
                            break;
                        case ".jpg":
                        case ".jpeg":
                            chart_Main.SaveImage(filePath, ChartImageFormat.Jpeg);
                            break;
                        //case ".bmp":
                        //    chart_Main.SaveImage(filePath, ChartImageFormat.Bmp);
                        //    break;
                        //case ".pdf":
                        //    // 需要额外的库来支持PDF导出，这里使用iTextSharp作为示例
                        //    ExportChartToPdf(filePath);
                        //    break;
                        default:
                            // 默认保存为PNG
                            chart_Main.SaveImage(filePath, ChartImageFormat.Png);
                            break;
                    }

                    MessageBox.Show($"图表已成功保存到: {filePath}", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存图表时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 使用iTextSharp导出PDF（需要添加NuGet包iTextSharp）
        //private void ExportChartToPdf(string filePath)
        //{
        //    // 临时保存为图片
        //    string tempImagePath = Path.GetTempFileName() + ".png";
        //    chart_Main.SaveImage(tempImagePath, ChartImageFormat.Png);

        //    try
        //    {
        //        // 创建PDF文档
        //        using (var fs = new FileStream(filePath, FileMode.Create))
        //        {
        //            using (var document = new iTextSharp.text.Document())
        //            {
        //                iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
        //                document.Open();

        //                // 添加图表图片到PDF
        //                var image = iTextSharp.text.Image.GetInstance(tempImagePath);
        //                image.ScaleToFit(document.PageSize.Width - document.LeftMargin - document.RightMargin,
        //                                document.PageSize.Height - document.TopMargin - document.BottomMargin);
        //                document.Add(image);
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        // 删除临时图片文件
        //        if (File.Exists(tempImagePath))
        //        {
        //            File.Delete(tempImagePath);
        //        }
        //    }
        //}
        #endregion

        #region 重置缩放
        private void btn_ResetZoom_Click(object sender, EventArgs e)
        {
            chart_Main.ChartAreas[0].AxisX.ScaleView.ZoomReset();
            chart_Main.ChartAreas[0].AxisY.ScaleView.ZoomReset();
        }
        #endregion

        #region  拖拽
        private void FrmMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // 检查文件扩展名是否是.csv或.logfile（不区分大小写）
                string fileExt = Path.GetExtension(files[0]).ToLower();
                if (fileExt == ".csv" || fileExt == ".logfile")
                {
                    e.Effect = DragDropEffects.Copy; // 显示复制图标
                    lbl_Status.Text = "松开鼠标加载文件"; // 可选：添加状态提示
                }
                else
                {
                    e.Effect = DragDropEffects.None; // 不支持的格式
                    lbl_Status.Text = "仅支持.csv和.logfile文件"; // 可选：添加状态提示
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void FrmMain_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string filePath = files[0];

                    // 更新UI显示文件路径
                    txtb_CSVFilePath.Text = filePath;

                    // 仅加载文件头（不自动加载数据）
                    LoadCSVHeaders(filePath);

                    // 提示用户文件已加载（可选）
                    lbl_Status.Text = $"已加载文件: {Path.GetFileName(filePath)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"拖放文件失败: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnDragEnter(DragEventArgs drgevent)
        {
            base.OnDragEnter(drgevent);
            this.BackColor = Color.LightBlue; // 拖入时改变背景色
        }

        protected override void OnDragLeave(EventArgs e)
        {
            base.OnDragLeave(e);
            this.BackColor = SystemColors.Control; // 恢复原背景色
        }

        protected override void OnDragDrop(DragEventArgs drgevent)
        {
            base.OnDragDrop(drgevent);
            this.BackColor = SystemColors.Control; // 恢复原背景色
        }

        #endregion
    }
}