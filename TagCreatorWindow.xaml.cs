using System;
using System.Windows;
using CAD_TagCreator.Services;

namespace CAD_TagCreator
{
    /// <summary>
    /// NodeCreatorWindow.xaml 的互動邏輯
    /// </summary>
    public partial class NodeCreatorWindow : Window
    {
        private TagCreatorService _service;

        public NodeCreatorWindow()
        {
            InitializeComponent();
            _service = new TagCreatorService(this);
        }

        /// <summary>
        /// 開始建立按鈕點擊事件
        /// </summary>
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            // 驗證輸入
            string prefix = TextBoxPrefix.Text.Trim();
            if (string.IsNullOrEmpty(prefix))
            {
                MessageBox.Show("請輸入代碼前綴！", "輸入錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(TextBoxStartNumber.Text.Trim(), out int startNumber) || startNumber <= 0)
            {
                MessageBox.Show("請輸入正確的起始號碼！", "輸入錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(TextBoxZoomRatio.Text.Trim(), out double zoomRatio) || zoomRatio <= 0)
            {
                MessageBox.Show("請輸入正確的標籤縮放比例！\n必須是大於 0 的數值，例如：300", "輸入錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 啟動節點建立流程（自動建立線段）
            _service.StartNodeCreation(prefix, startNumber, true, zoomRatio);

            // 啟用結束按鈕
            ExitButton.IsEnabled = true;
        }

        /// <summary>
        /// 結束並輸出按鈕點擊事件
        /// </summary>
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            _service.FinishAndExport();

            // 重置UI
            ExitButton.IsEnabled = false;
        }

        /// <summary>
        /// 更新統計資訊
        /// </summary>
        public void UpdateStatistics(int nodesCount, int linesCount, int areasCount)
        {
            TextNodesCreated.Text = nodesCount.ToString();
            TextLinesCreated.Text = linesCount.ToString();
            TextAreasCreated.Text = areasCount.ToString();
        }

        /// <summary>
        /// 更新起始號碼
        /// </summary>
        public void UpdateStartNumber(int startNumber)
        {
            TextBoxStartNumber.Text = startNumber.ToString();
        }

        /// <summary>
        /// 視窗關閉事件
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _service?.Cleanup();
        }
    }
}
