using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Media.Imaging;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

namespace CAD_TagCreator
{
    /// <summary>
    /// Ribbon 配置類，用於建立 Ribbon 按鈕
    /// </summary>
    public class RibbonConfig : IExtensionApplication
    {
        /// <summary>
        /// 插件初始化
        /// </summary>
        public void Initialize()
        {
            try
            {
                // 延遲執行，等待 Ribbon 完全載入
                Application.Idle += OnIdle;
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    $"\nRibbon 初始化錯誤: {ex.Message}\n");
            }
        }

        private void OnIdle(object sender, EventArgs e)
        {
            // 只執行一次
            Application.Idle -= OnIdle;

            try
            {
                CreateRibbonTab();
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    $"\n建立 Ribbon 錯誤: {ex.Message}\n");
            }
        }

        /// <summary>
        /// 建立 Ribbon 選項卡和按鈕
        /// </summary>
        private void CreateRibbonTab()
        {
            // 獲取 Ribbon 控件
            RibbonControl ribbonControl = ComponentManager.Ribbon;

            if (ribbonControl == null)
            {
                return;
            }

            // 查找或建立自訂選項卡
            RibbonTab tab = GetOrCreateTab(ribbonControl, "ID_NodeCreator_Tab", "中興軌一");

            // 建立面板
            RibbonPanel panel = GetOrCreatePanel(tab, "ID_NodeCreator_Panel", "節點建立工具");

            // 建立節點建立工具按鈕
            CreateNodeCreatorButton(panel);
        }

        /// <summary>
        /// 獲取或建立 Ribbon Tab
        /// </summary>
        private RibbonTab GetOrCreateTab(RibbonControl ribbonControl, string tabId, string tabTitle)
        {
            // 查找現有 Tab
            foreach (RibbonTab tab in ribbonControl.Tabs)
            {
                if (tab.Id == tabId)
                {
                    return tab;
                }
            }

            // 建立新 Tab
            RibbonTab newTab = new RibbonTab
            {
                Title = tabTitle,
                Id = tabId
            };

            ribbonControl.Tabs.Add(newTab);
            return newTab;
        }

        /// <summary>
        /// 獲取或建立 Ribbon Panel
        /// </summary>
        private RibbonPanel GetOrCreatePanel(RibbonTab tab, string panelId, string panelTitle)
        {
            // 查找現有 Panel
            foreach (RibbonPanel panel in tab.Panels)
            {
                if (panel.Source.Id == panelId)
                {
                    return panel;
                }
            }

            // 建立新 Panel
            RibbonPanelSource panelSource = new RibbonPanelSource
            {
                Title = panelTitle,
                Id = panelId
            };

            RibbonPanel newPanel = new RibbonPanel
            {
                Source = panelSource
            };

            tab.Panels.Add(newPanel);
            return newPanel;
        }

        /// <summary>
        /// 建立節點建立工具按鈕
        /// </summary>
        private void CreateNodeCreatorButton(RibbonPanel panel)
        {
            // 建立按鈕
            RibbonButton button = new RibbonButton
            {
                Text = "節點建立工具",
                ShowText = false,  // 不顯示文字
                ShowImage = true,
                Id = "ID_NodeCreator_Button",
                // 按鈕大小
                Size = RibbonItemSize.Large,
                // 設定工具提示
                ToolTip = "節點建立工具",
                // 命令處理器
                CommandHandler = new NodeCreatorCommandHandler()
            };

            // 設定按鈕圖標（使用相同圖標，你可以之後更換）
            try
            {
                button.Image = LoadIconFromResources(16);      // 16x16 小圖標
                button.LargeImage = LoadIconFromResources(32); // 32x32 大圖標
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    $"\n載入圖標錯誤: {ex.Message}\n");
            }

            // 添加按鈕到面板
            panel.Source.Items.Add(button);
        }

        /// <summary>
        /// 從資源載入並縮放圖標
        /// </summary>
        private BitmapImage LoadIconFromResources(int size)
        {
            // 從資源載入圖片
            Bitmap originalBitmap = Properties.Resources.nodes;

            if (originalBitmap == null)
            {
                throw new System.Exception("無法載入 nodes 圖標資源");
            }

            // 創建指定尺寸的新 Bitmap
            using (Bitmap resizedBitmap = new Bitmap(size, size))
            {
                using (Graphics g = Graphics.FromImage(resizedBitmap))
                {
                    // 使用高品質縮放
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.Clear(Color.Transparent);

                    // 繪製縮放後的圖片
                    g.DrawImage(originalBitmap, 0, 0, size, size);
                }

                // 將 Bitmap 轉換為 BitmapImage
                return ConvertBitmapToBitmapImage(resizedBitmap);
            }
        }

        /// <summary>
        /// 將 Bitmap 轉換為 BitmapImage
        /// </summary>
        private BitmapImage ConvertBitmapToBitmapImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();

                return bitmapImage;
            }
        }

        /// <summary>
        /// 插件終止
        /// </summary>
        public void Terminate()
        {
            // 清理資源
        }
    }

    /// <summary>
    /// 節點建立工具 Ribbon 命令處理器
    /// </summary>
    public class NodeCreatorCommandHandler : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            try
            {
                // 獲取當前文檔
                Document doc = Application.DocumentManager.MdiActiveDocument;

                if (doc == null)
                {
                    return;
                }

                // 直接建立並顯示 WPF 視窗
                NodeCreatorWindow window = new NodeCreatorWindow();

                // 使用 AutoCAD 的視窗作為父視窗
                Application.ShowModelessWindow(window);
            }
            catch (System.Exception ex)
            {
                // 如果出錯，顯示錯誤訊息
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    $"\n錯誤: {ex.Message}\n");
            }
        }
    }
}
