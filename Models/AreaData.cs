using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Linq;
using System;

namespace CAD_TagCreator.Models
{
    /// <summary>
    /// 面域資料模型
    /// </summary>
    public class AreaData
    {
        /// <summary>
        /// 面域標籤 (例如: A-ABC-0001)
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 中心點 X 座標
        /// </summary>
        public double CenterX { get; set; }

        /// <summary>
        /// 中心點 Y 座標
        /// </summary>
        public double CenterY { get; set; }

        /// <summary>
        /// 頂點數量
        /// </summary>
        public int VertexCount { get; set; }

        /// <summary>
        /// 圖層名稱
        /// </summary>
        public string LayerName { get; set; }

        /// <summary>
        /// 節點標籤列表
        /// </summary>
        public List<string> NodeLabels { get; set; }

        /// <summary>
        /// 面積
        /// </summary>
        public double Area { get; set; }

        /// <summary>
        /// 中心點位置
        /// </summary>
        public Point3d CenterPoint
        {
            get { return new Point3d(CenterX, CenterY, 0); }
        }

        /// <summary>
        /// 節點標籤字串（用逗號分隔）
        /// </summary>
        public string NodeLabelsString
        {
            get { return string.Join(", ", NodeLabels); }
        }

        /// <summary>
        /// 建構函數
        /// </summary>
        public AreaData(string label, Point3d centerPoint, List<NodeData> nodes, string layerName, double area)
        {
            Label = label;
            CenterX = centerPoint.X;
            CenterY = centerPoint.Y;
            VertexCount = nodes.Count;
            LayerName = layerName;
            NodeLabels = nodes.Select(n => n.Label).ToList();
            Area = area;
        }

        /// <summary>
        /// 計算多邊形面積（使用 Shoelace 公式）
        /// </summary>
        public static double CalculatePolygonArea(List<NodeData> nodes)
        {
            if (nodes.Count < 3)
                return 0;

            double area = 0;
            for (int i = 0; i < nodes.Count; i++)
            {
                int j = (i + 1) % nodes.Count;
                area += nodes[i].X * nodes[j].Y;
                area -= nodes[j].X * nodes[i].Y;
            }

            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// 計算多邊形中心點
        /// </summary>
        public static Point3d CalculateCentroid(List<NodeData> nodes)
        {
            if (nodes.Count == 0)
                return Point3d.Origin;

            double sumX = 0, sumY = 0;
            foreach (var node in nodes)
            {
                sumX += node.X;
                sumY += node.Y;
            }

            return new Point3d(sumX / nodes.Count, sumY / nodes.Count, 0);
        }

        public override string ToString()
        {
            return $"{Label} ({CenterX:F3}, {CenterY:F3}) Area: {Area:F3} Vertices: {VertexCount}";
        }
    }
}
