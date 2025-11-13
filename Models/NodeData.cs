using Autodesk.AutoCAD.Geometry;

namespace CAD_TagCreator.Models
{
    /// <summary>
    /// 節點資料模型
    /// </summary>
    public class NodeData
    {
        /// <summary>
        /// 節點標籤 (例如: P-ABC-0001)
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// X 座標
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// Y 座標
        /// </summary>
        public double Y { get; set; }

        /// <summary>
        /// 圖層名稱
        /// </summary>
        public string LayerName { get; set; }

        /// <summary>
        /// 位置點
        /// </summary>
        public Point3d Position
        {
            get { return new Point3d(X, Y, 0); }
        }

        /// <summary>
        /// 建構函數
        /// </summary>
        public NodeData(string label, double x, double y, string layerName)
        {
            Label = label;
            X = x;
            Y = y;
            LayerName = layerName;
        }

        /// <summary>
        /// 從 Point3d 建立 NodeData
        /// </summary>
        public NodeData(string label, Point3d position, string layerName)
        {
            Label = label;
            X = position.X;
            Y = position.Y;
            LayerName = layerName;
        }

        public override string ToString()
        {
            return $"{Label} ({X:F3}, {Y:F3})";
        }
    }
}
