using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

namespace CAD_TagCreator
{
    /// <summary>
    /// AutoCAD 命令類
    /// </summary>
    public class Commands
    {
        /// <summary>
        /// NODECREATE 命令 - 打開節點建立工具視窗
        /// </summary>
        [CommandMethod("NODECREATE")]
        public void NodeCreate()
        {
            try
            {
                // 獲取當前文檔
                Document doc = Application.DocumentManager.MdiActiveDocument;

                if (doc == null)
                {
                    return;
                }

                // 建立並顯示 WPF 視窗
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
