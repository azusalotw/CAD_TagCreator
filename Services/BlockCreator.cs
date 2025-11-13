using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using CAD_TagCreator.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CAD_TagCreator.Services
{
    /// <summary>
    /// 圖塊建立器
    /// </summary>
    public class BlockCreator
    {
        private const double DUPLICATE_POINT_TOLERANCE = 1.0;

        private double _zoomRatio = 300.0;
        private double NODE_CIRCLE_RADIUS => 1 * _zoomRatio;
        private double ATTRIBUTE_TEXT_HEIGHT => 2 * _zoomRatio;
        private double LINE_ATTRIBUTE_TEXT_HEIGHT => 2 * _zoomRatio;
        private double AREA_ATTRIBUTE_TEXT_HEIGHT => 2 * _zoomRatio;

        private Document _doc;
        private Database _db;

        public BlockCreator(Document doc)
        {
            _doc = doc;
            _db = doc.Database;
        }

        /// <summary>
        /// 設定縮放比例
        /// </summary>
        public void SetZoomRatio(double zoomRatio)
        {
            _zoomRatio = zoomRatio;
        }

        /// <summary>
        /// 建立節點圖塊定義
        /// </summary>
        public ObjectId CreateNodeBlockDefinition(string blockName, string prefix, int nodeIndex)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                // 加入圖塊表
                bt.UpgradeOpen();
                ObjectId blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // 建立圓形
                Circle circle = new Circle();
                circle.Center = Point3d.Origin;
                circle.Radius = NODE_CIRCLE_RADIUS;
                btr.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);

                // 建立垂直線
                Line vLine = new Line(
                    new Point3d(0, -NODE_CIRCLE_RADIUS * 1.2, 0),
                    new Point3d(0, NODE_CIRCLE_RADIUS * 1.2, 0)
                );
                btr.AppendEntity(vLine);
                tr.AddNewlyCreatedDBObject(vLine, true);

                // 建立水平線
                Line hLine = new Line(
                    new Point3d(-NODE_CIRCLE_RADIUS * 1.2, 0, 0),
                    new Point3d(NODE_CIRCLE_RADIUS * 1.2, 0, 0)
                );
                btr.AppendEntity(hLine);
                tr.AddNewlyCreatedDBObject(hLine, true);

                // 建立屬性定義
                AttributeDefinition attDef = new AttributeDefinition();
                attDef.Position = new Point3d(NODE_CIRCLE_RADIUS, NODE_CIRCLE_RADIUS, 0);
                attDef.Height = ATTRIBUTE_TEXT_HEIGHT;
                attDef.Tag = "LABEL";
                attDef.Prompt = "LABEL";
                attDef.TextString = $"P-{prefix}-{nodeIndex:D4}";
                attDef.VerticalMode = TextVerticalMode.TextBase;
                attDef.HorizontalMode = TextHorizontalMode.TextLeft;

                btr.AppendEntity(attDef);
                tr.AddNewlyCreatedDBObject(attDef, true);

                tr.Commit();
                return blockId;
            }
        }

        /// <summary>
        /// 插入節點圖塊
        /// </summary>
        public ObjectId InsertNodeBlock(Point3d position, string blockName, string layerName, string label)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // 插入圖塊參照
                BlockReference blockRef = new BlockReference(position, bt[blockName]);
                blockRef.Layer = layerName;

                ObjectId blockRefId = modelSpace.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);

                // 設定屬性
                BlockTableRecord btr = tr.GetObject(bt[blockName], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId id in btr)
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                    if (obj is AttributeDefinition attDef)
                    {
                        AttributeReference attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                        attRef.TextString = label;
                        attRef.Layer = layerName;

                        blockRef.AttributeCollection.AppendAttribute(attRef);
                        tr.AddNewlyCreatedDBObject(attRef, true);
                    }
                }

                tr.Commit();
                return blockRefId;
            }
        }

        /// <summary>
        /// 建立線段圖塊定義
        /// </summary>
        public ObjectId CreateLineBlockDefinition(string blockName, Point3d startPoint, Point3d endPoint,
            Point3d midPoint, string prefix, int lineIndex)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                // 加入圖塊表
                bt.UpgradeOpen();
                ObjectId blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // 計算相對於中點的座標
                Point3d relativeStart = new Point3d(
                    startPoint.X - midPoint.X,
                    startPoint.Y - midPoint.Y,
                    0
                );
                Point3d relativeEnd = new Point3d(
                    endPoint.X - midPoint.X,
                    endPoint.Y - midPoint.Y,
                    0
                );

                // 建立線段
                Line line = new Line(relativeStart, relativeEnd);
                btr.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);

                // 建立屬性定義
                AttributeDefinition attDef = new AttributeDefinition();
                attDef.Position = new Point3d(0, 2, 0);
                attDef.Height = LINE_ATTRIBUTE_TEXT_HEIGHT;
                attDef.Tag = "LABEL";
                attDef.Prompt = "LABEL";
                attDef.TextString = $"L-{prefix}-{lineIndex:D4}";
                attDef.VerticalMode = TextVerticalMode.TextBase;
                attDef.HorizontalMode = TextHorizontalMode.TextLeft;

                btr.AppendEntity(attDef);
                tr.AddNewlyCreatedDBObject(attDef, true);

                tr.Commit();
                return blockId;
            }
        }

        /// <summary>
        /// 插入線段圖塊
        /// </summary>
        public ObjectId InsertLineBlock(Point3d midPoint, string blockName, string layerName, string label)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // 插入圖塊參照
                BlockReference blockRef = new BlockReference(midPoint, bt[blockName]);
                blockRef.Layer = layerName;

                ObjectId blockRefId = modelSpace.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);

                // 設定屬性
                BlockTableRecord btr = tr.GetObject(bt[blockName], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId id in btr)
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                    if (obj is AttributeDefinition attDef)
                    {
                        AttributeReference attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                        attRef.TextString = label;
                        attRef.Layer = layerName;

                        blockRef.AttributeCollection.AppendAttribute(attRef);
                        tr.AddNewlyCreatedDBObject(attRef, true);
                    }
                }

                tr.Commit();
                return blockRefId;
            }
        }

        /// <summary>
        /// 建立面域圖塊定義
        /// </summary>
        public ObjectId CreateAreaBlockDefinition(string blockName, List<NodeData> nodes,
            Point3d centroid, string areaLabel)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                // 加入圖塊表
                bt.UpgradeOpen();
                ObjectId blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // 建立多段線（相對於中心點）
                Polyline pline = new Polyline();
                for (int i = 0; i < nodes.Count; i++)
                {
                    Point2d pt = new Point2d(
                        nodes[i].X - centroid.X,
                        nodes[i].Y - centroid.Y
                    );
                    pline.AddVertexAt(i, pt, 0, 0, 0);
                }
                pline.Closed = true;
                pline.LineWeight = LineWeight.LineWeight025;

                btr.AppendEntity(pline);
                tr.AddNewlyCreatedDBObject(pline, true);

                // 建立屬性定義
                AttributeDefinition attDef = new AttributeDefinition();
                attDef.Position = new Point3d(0, 5, 0);
                attDef.Height = AREA_ATTRIBUTE_TEXT_HEIGHT;
                attDef.Tag = "LABEL";
                attDef.Prompt = "LABEL";
                attDef.TextString = areaLabel;
                attDef.VerticalMode = TextVerticalMode.TextBase;
                attDef.HorizontalMode = TextHorizontalMode.TextLeft;

                btr.AppendEntity(attDef);
                tr.AddNewlyCreatedDBObject(attDef, true);

                tr.Commit();
                return blockId;
            }
        }

        /// <summary>
        /// 插入面域圖塊
        /// </summary>
        public ObjectId InsertAreaBlock(Point3d centroid, string blockName, string layerName, string label)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(_db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // 插入圖塊參照
                BlockReference blockRef = new BlockReference(centroid, bt[blockName]);
                blockRef.Layer = layerName;

                ObjectId blockRefId = modelSpace.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);

                // 設定屬性
                BlockTableRecord btr = tr.GetObject(bt[blockName], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId id in btr)
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                    if (obj is AttributeDefinition attDef)
                    {
                        AttributeReference attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                        attRef.TextString = label;
                        attRef.Layer = layerName;

                        blockRef.AttributeCollection.AppendAttribute(attRef);
                        tr.AddNewlyCreatedDBObject(attRef, true);
                    }
                }

                tr.Commit();
                return blockRefId;
            }
        }

        /// <summary>
        /// 建立或取得圖層
        /// </summary>
        public ObjectId CreateOrGetLayer(string layerName)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                LayerTable lt = tr.GetObject(_db.LayerTableId, OpenMode.ForRead) as LayerTable;

                if (lt.Has(layerName))
                {
                    ObjectId layerId = lt[layerName];
                    tr.Commit();
                    return layerId;
                }

                // 建立新圖層
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                    Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1); // 紅色
                ltr.IsPlottable = false;

                ObjectId newLayerId = lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);

                tr.Commit();
                return newLayerId;
            }
        }

        /// <summary>
        /// 刪除實體
        /// </summary>
        public void DeleteEntity(ObjectId entityId)
        {
            if (entityId.IsNull)
                return;

            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                DBObject obj = tr.GetObject(entityId, OpenMode.ForWrite);
                obj.Erase();
                tr.Commit();
            }
        }

        /// <summary>
        /// 檢查點是否重複
        /// </summary>
        public static bool IsPointDuplicate(Point3d point1, Point3d point2)
        {
            double distance = point1.DistanceTo(point2);
            return distance < DUPLICATE_POINT_TOLERANCE;
        }
    }
}
