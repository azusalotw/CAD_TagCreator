using CAD_TagCreator.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CAD_TagCreator.Services
{
    /// <summary>
    /// Excel 輸出器
    /// </summary>
    public class ExcelExporter
    {
        /// <summary>
        /// 輸出資料到 Excel
        /// </summary>
        public string ExportToExcel(List<NodeData> nodes, List<LineData> lines, List<AreaData> areas, string drawingName)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                // 建立 Excel 應用程式
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();

                // 建立節點工作表
                if (nodes.Count > 0)
                {
                    CreateNodeWorksheet(workbook, nodes);
                }

                // 建立線段工作表
                if (lines.Count > 0)
                {
                    CreateLineWorksheet(workbook, lines, nodes);
                }

                // 建立面域工作表
                if (areas.Count > 0)
                {
                    CreateAreaWorksheet(workbook, areas);
                }

                // 儲存檔案
                string fileName = SaveWorkbook(workbook, drawingName);
                return fileName;
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Excel 輸出錯誤: {ex.Message}", ex);
            }
            finally
            {
                // 釋放 COM 物件
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        /// <summary>
        /// 建立節點工作表
        /// </summary>
        private void CreateNodeWorksheet(Excel.Workbook workbook, List<NodeData> nodes)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            worksheet.Name = "節點清單";

            // 設定標題
            worksheet.Cells[1, 1] = "節點標籤";
            worksheet.Cells[1, 2] = "X座標";
            worksheet.Cells[1, 3] = "Y座標";
            worksheet.Cells[1, 4] = "圖層名稱";

            // 格式化標題
            Excel.Range headerRange = worksheet.Range["A1", "D1"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            headerRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            // 填入資料
            for (int i = 0; i < nodes.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = nodes[i].Label;
                worksheet.Cells[i + 2, 2] = Math.Round(nodes[i].X, 3);
                worksheet.Cells[i + 2, 3] = Math.Round(nodes[i].Y, 3);
                worksheet.Cells[i + 2, 4] = nodes[i].LayerName;
            }

            // 自動調整欄寬
            worksheet.Columns.AutoFit();
        }

        /// <summary>
        /// 建立線段工作表
        /// </summary>
        private void CreateLineWorksheet(Excel.Workbook workbook, List<LineData> lines, List<NodeData> nodes)
        {
            Excel.Worksheet worksheet = workbook.Worksheets.Add();
            worksheet.Name = "線段清單";

            // 設定標題
            worksheet.Cells[1, 1] = "線段標籤";
            worksheet.Cells[1, 2] = "起點節點標籤";
            worksheet.Cells[1, 3] = "終點節點標籤";
            worksheet.Cells[1, 4] = "起點X";
            worksheet.Cells[1, 5] = "起點Y";
            worksheet.Cells[1, 6] = "終點X";
            worksheet.Cells[1, 7] = "終點Y";
            worksheet.Cells[1, 8] = "長度";
            worksheet.Cells[1, 9] = "角度(度)";
            worksheet.Cells[1, 10] = "圖層名稱";

            // 格式化標題
            Excel.Range headerRange = worksheet.Range["A1", "J1"];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            headerRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            // 填入資料
            for (int i = 0; i < lines.Count; i++)
            {
                string startNodeLabel = GetNodeLabelAtPoint(lines[i].StartPoint, nodes, ExtractPrefix(lines[i].Label));
                string endNodeLabel = GetNodeLabelAtPoint(lines[i].EndPoint, nodes, ExtractPrefix(lines[i].Label));

                worksheet.Cells[i + 2, 1] = lines[i].Label;
                worksheet.Cells[i + 2, 2] = startNodeLabel;
                worksheet.Cells[i + 2, 3] = endNodeLabel;
                worksheet.Cells[i + 2, 4] = Math.Round(lines[i].StartX, 3);
                worksheet.Cells[i + 2, 5] = Math.Round(lines[i].StartY, 3);
                worksheet.Cells[i + 2, 6] = Math.Round(lines[i].EndX, 3);
                worksheet.Cells[i + 2, 7] = Math.Round(lines[i].EndY, 3);
                worksheet.Cells[i + 2, 8] = Math.Round(lines[i].Length, 3);
                worksheet.Cells[i + 2, 9] = Math.Round(lines[i].AngleDegrees, 2);
                worksheet.Cells[i + 2, 10] = lines[i].LayerName;
            }

            // 自動調整欄寬
            worksheet.Columns.AutoFit();
        }

        /// <summary>
        /// 建立面域工作表
        /// </summary>
        private void CreateAreaWorksheet(Excel.Workbook workbook, List<AreaData> areas)
        {
            Excel.Worksheet worksheet = workbook.Worksheets.Add();
            worksheet.Name = "面域清單";

            // 找出最多節點數量
            int maxNodes = areas.Max(a => a.VertexCount);

            // 設定標題
            worksheet.Cells[1, 1] = "面域標籤";
            worksheet.Cells[1, 2] = "中心點X";
            worksheet.Cells[1, 3] = "中心點Y";
            worksheet.Cells[1, 4] = "頂點數量";
            worksheet.Cells[1, 5] = "圖層名稱";
            worksheet.Cells[1, 6] = "面積";

            // 動態建立節點欄位標題
            for (int i = 1; i <= maxNodes; i++)
            {
                worksheet.Cells[1, 6 + i] = $"節點{i}";
            }

            // 格式化標題
            Excel.Range headerRange = worksheet.Range[
                worksheet.Cells[1, 1],
                worksheet.Cells[1, 6 + maxNodes]
            ];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            headerRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            // 填入資料
            for (int i = 0; i < areas.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = areas[i].Label;
                worksheet.Cells[i + 2, 2] = Math.Round(areas[i].CenterX, 3);
                worksheet.Cells[i + 2, 3] = Math.Round(areas[i].CenterY, 3);
                worksheet.Cells[i + 2, 4] = areas[i].VertexCount;
                worksheet.Cells[i + 2, 5] = areas[i].LayerName;
                worksheet.Cells[i + 2, 6] = Math.Round(areas[i].Area, 3);

                // 填入節點標籤
                for (int j = 0; j < areas[i].NodeLabels.Count; j++)
                {
                    worksheet.Cells[i + 2, 7 + j] = areas[i].NodeLabels[j];
                }
            }

            // 自動調整欄寬
            worksheet.Columns.AutoFit();
        }

        /// <summary>
        /// 儲存工作簿
        /// </summary>
        private string SaveWorkbook(Excel.Workbook workbook, string drawingName)
        {
            string directory = Path.GetDirectoryName(drawingName);
            if (string.IsNullOrEmpty(directory))
            {
                directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            string fileName = Path.Combine(
                directory,
                $"節點線段清單_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            );

            workbook.SaveAs(fileName);
            return fileName;
        }

        /// <summary>
        /// 取得點位上的節點標籤
        /// </summary>
        private string GetNodeLabelAtPoint(Autodesk.AutoCAD.Geometry.Point3d point, List<NodeData> nodes, string prefix)
        {
            const double tolerance = 0.001;

            var node = nodes.FirstOrDefault(n =>
            {
                if (ExtractPrefix(n.Label) != prefix)
                    return false;

                double distance = Math.Sqrt(
                    Math.Pow(point.X - n.X, 2) +
                    Math.Pow(point.Y - n.Y, 2)
                );
                return distance < tolerance;
            });

            return node?.Label ?? "";
        }

        /// <summary>
        /// 從標籤中提取前綴
        /// </summary>
        private string ExtractPrefix(string label)
        {
            var parts = label.Split('-');
            if (parts.Length >= 3)
                return parts[1];
            return "";
        }
    }
}
