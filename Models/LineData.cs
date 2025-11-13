using Autodesk.AutoCAD.Geometry;
using System;

namespace CAD_TagCreator.Models
{
    /// <summary>
    /// 線段資料模型
    /// </summary>
    public class LineData
    {
        /// <summary>
        /// 線段標籤 (例如: L-ABC-0001)
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 起點 X 座標
        /// </summary>
        public double StartX { get; set; }

        /// <summary>
        /// 起點 Y 座標
        /// </summary>
        public double StartY { get; set; }

        /// <summary>
        /// 終點 X 座標
        /// </summary>
        public double EndX { get; set; }

        /// <summary>
        /// 終點 Y 座標
        /// </summary>
        public double EndY { get; set; }

        /// <summary>
        /// 線段長度
        /// </summary>
        public double Length { get; set; }

        /// <summary>
        /// 線段角度（度）
        /// </summary>
        public double AngleDegrees { get; set; }

        /// <summary>
        /// 圖層名稱
        /// </summary>
        public string LayerName { get; set; }

        /// <summary>
        /// 起點位置
        /// </summary>
        public Point3d StartPoint
        {
            get { return new Point3d(StartX, StartY, 0); }
        }

        /// <summary>
        /// 終點位置
        /// </summary>
        public Point3d EndPoint
        {
            get { return new Point3d(EndX, EndY, 0); }
        }

        /// <summary>
        /// 中點位置
        /// </summary>
        public Point3d MidPoint
        {
            get { return new Point3d((StartX + EndX) / 2, (StartY + EndY) / 2, 0); }
        }

        /// <summary>
        /// 建構函數
        /// </summary>
        public LineData(string label, Point3d startPoint, Point3d endPoint, string layerName)
        {
            Label = label;
            StartX = startPoint.X;
            StartY = startPoint.Y;
            EndX = endPoint.X;
            EndY = endPoint.Y;
            LayerName = layerName;

            // 計算長度和角度
            CalculateLengthAndAngle();
        }

        /// <summary>
        /// 計算長度和角度
        /// </summary>
        private void CalculateLengthAndAngle()
        {
            double deltaX = EndX - StartX;
            double deltaY = EndY - StartY;
            Length = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

            // 計算角度（弧度轉度）
            double angleRad;
            if (deltaX == 0)
            {
                angleRad = deltaY > 0 ? Math.PI / 2 : -Math.PI / 2;
            }
            else
            {
                angleRad = Math.Atan(deltaY / deltaX);
                if (deltaX < 0)
                {
                    angleRad += Math.PI;
                }
                else if (deltaY < 0)
                {
                    angleRad += 2 * Math.PI;
                }
            }

            AngleDegrees = angleRad * 180.0 / Math.PI;
        }

        public override string ToString()
        {
            return $"{Label} ({StartX:F3},{StartY:F3}) -> ({EndX:F3},{EndY:F3}) Length: {Length:F3}";
        }
    }
}
