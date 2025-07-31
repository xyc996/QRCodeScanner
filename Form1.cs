using System;
using System.Drawing;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using ZXing;

namespace QRCodeScanner
{
    public partial class Form1 : Form
    {
        private VideoCaptureDevice videoSource;
        private BarcodeReader barcodeReader;
        private Label statusLabel;
        private PictureBox videoBox;
        // 新增：用于存储当前显示的图像，方便释放
        private Bitmap currentFrame;

        public Form1()
        {
            InitializeComponent();
            InitializeUI();
            barcodeReader = new BarcodeReader();
            barcodeReader.Options.PossibleFormats = new[] { BarcodeFormat.QR_CODE };
            barcodeReader.Options.TryHarder = true;
        }

        private void InitializeUI()
        {
            this.Text = "二维码扫描打印工具";
            this.Width = 800;
            this.Height = 600;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // 状态标签
            statusLabel = new Label { Text = "状态：等待启动扫描", Location = new Point(20, 20), Size = new Size(740, 20), ForeColor = Color.DarkBlue };
            this.Controls.Add(statusLabel);

            // 开始/停止按钮
            Button startBtn = new Button { Text = "开始扫描", Location = new Point(20, 50), Size = new Size(100, 30) };
            startBtn.Click += StartScan_Click;
            this.Controls.Add(startBtn);

            Button stopBtn = new Button { Text = "停止扫描", Location = new Point(130, 50), Size = new Size(100, 30), Enabled = false };
            stopBtn.Click += StopScan_Click;
            this.Controls.Add(stopBtn);

            // 摄像头画面控件
            videoBox = new PictureBox { Location = new Point(20, 90), Size = new Size(740, 400), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.Black, SizeMode = PictureBoxSizeMode.StretchImage };
            this.Controls.Add(videoBox);

            // 提示标签
            Label infoLabel = new Label { Text = "提示：对准二维码自动识别，支持网页链接", Location = new Point(20, 500), Size = new Size(740, 20), ForeColor = Color.Gray };
            this.Controls.Add(infoLabel);
        }

        private void StartScan_Click(object sender, EventArgs e)
        {
            try
            {
                StopScan(); // 先停止可能的旧实例
                FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                if (videoDevices.Count == 0)
                {
                    UpdateStatus("错误：未检测到摄像头", Color.Red);
                    return;
                }

                videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
                videoSource.NewFrame += VideoSource_NewFrame;
                videoSource.Start();
                UpdateStatus($"状态：启动成功（{videoDevices[0].Name}）", Color.Green);
                ((Button)sender).Enabled = false;
                this.Controls[2].Enabled = true;
            }
            catch (Exception ex)
            {
                UpdateStatus($"启动失败：{ex.Message}", Color.Red);
            }
        }

        private void StopScan_Click(object sender, EventArgs e)
        {
            StopScan();
            UpdateStatus("状态：已停止", Color.DarkBlue);
            ((Button)sender).Enabled = false;
            this.Controls[1].Enabled = true;
        }

        private void StopScan()
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.NewFrame -= VideoSource_NewFrame;
                videoSource.SignalToStop();
                videoSource = null;
            }
            // 释放当前图像资源
            if (currentFrame != null)
            {
                videoBox.Image = null;
                currentFrame.Dispose(); // 关键：释放锁定的Bitmap
                currentFrame = null;
            }
        }

        // 修复核心：正确处理Bitmap锁定和释放
        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                // 1. 克隆新帧（避免直接使用原始帧，防止被设备锁定）
                using (Bitmap newFrame = (Bitmap)eventArgs.Frame.Clone())
                {
                    // 2. 先释放旧图像资源（关键：避免锁定冲突）
                    Bitmap oldFrame = currentFrame;
                    currentFrame = (Bitmap)newFrame.Clone(); // 再次克隆，确保新实例独立

                    // 3. 跨线程更新UI，显示新帧
                    videoBox.Invoke((Action)(() =>
                    {
                        videoBox.Image?.Dispose(); // 释放PictureBox当前显示的图像
                        videoBox.Image = currentFrame;
                    }));

                    // 4. 释放旧帧（延迟释放，确保UI已更新）
                    if (oldFrame != null)
                    {
                        oldFrame.Dispose();
                    }

                    // 5. 识别二维码（使用newFrame的克隆，避免干扰UI显示）
                    using (Bitmap decodeFrame = (Bitmap)newFrame.Clone())
                    {
                        Result result = barcodeReader.Decode(decodeFrame);
                        if (result != null && !string.IsNullOrEmpty(result.Text))
                        {
                            ProcessQRCode(result.Text);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"画面错误：{ex.Message}", Color.Orange);
            }
        }

        private void ProcessQRCode(string content)
        {
            StopScan();
            UpdateStatus($"已识别：{content}", Color.Purple);

            if (content.StartsWith("http://") || content.StartsWith("https://"))
            {
                if (MessageBox.Show($"打开并打印此链接？\n{content}", "识别结果", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    OpenAndPrintUrl(content);
                }
            }
            else
            {
                MessageBox.Show($"识别内容：\n{content}", "非网页链接");
            }

            this.Controls[1].Enabled = true;
            this.Controls[2].Enabled = false;
        }

        private void OpenAndPrintUrl(string url)
        {
            try
            {
                UpdateStatus($"打开网页：{url}", Color.Blue);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                System.Threading.Thread.Sleep(1800);
                SendKeys.SendWait("^p");
                UpdateStatus("已触发打印，请确认", Color.Green);
            }
            catch (Exception ex)
            {
                UpdateStatus($"打开失败：{ex.Message}", Color.Red);
            }
        }

        private void UpdateStatus(string text, Color color)
        {
            statusLabel.Invoke((Action)(() =>
            {
                statusLabel.Text = text;
                statusLabel.ForeColor = color;
            }));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            StopScan();
            base.OnFormClosing(e);
        }
    }
}