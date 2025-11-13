using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using CAD_TagCreator.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CAD_TagCreator.Services
{
    /// <summary>
    /// 節點建立服務
    /// </summary>
    public class TagCreatorService
    {
        private Document _doc;
        private Editor _ed;
        private Database _db;
        private BlockCreator _blockCreator;
        private NodeCreatorWindow _window;

        // 資料收集
        private List<NodeData> _allNodes;
        private List<LineData> _allLines;
        private List<AreaData> _allAreas;
        private List<NodeData> _currentSessionNodes;

        // 會話狀態
        private bool _isSessionActive;
        private string _currentPrefix;
        private int _currentNodeIndex;
        private int _currentLineIndex;
        private bool _shouldCreateLines;
        private string _currentLayerName;
        private ObjectId _targetLayerId;
        private Point3d? _previousNodePoint;
        private string _originalLayerName;

        // 物件 ID 追蹤（用於復原）
        private Dictionary<string, ObjectId> _nodeObjectIds;
        private Dictionary<string, ObjectId> _lineObjectIds;
        private Dictionary<string, ObjectId> _areaObjectIds;

        public TagCreatorService(NodeCreatorWindow window)
        {
            _window = window;
            _doc = Application.DocumentManager.MdiActiveDocument;
            _ed = _doc.Editor;
            _db = _doc.Database;
            _blockCreator = new BlockCreator(_doc);

            _allNodes = new List<NodeData>();
            _allLines = new List<LineData>();
            _allAreas = new List<AreaData>();
            _currentSessionNodes = new List<NodeData>();

            _nodeObjectIds = new Dictionary<string, ObjectId>();
            _lineObjectIds = new Dictionary<string, ObjectId>();
            _areaObjectIds = new Dictionary<string, ObjectId>();

            _isSessionActive = false;
        }

        /// <summary>
        /// 開始節點建立
        /// </summary>
        public void StartNodeCreation(string prefix, int startNumber, bool createLines, double zoomRatio)
        {
            try
            {
                // 更新 BlockCreator 的縮放比例
                _blockCreator.SetZoomRatio(zoomRatio);

                // 初始化或繼續會話
                if (!_isSessionActive)
                {
                    _isSessionActive = true;
                    _originalLayerName = _db.Clayer.ToString();
                    _currentNodeIndex = startNumber;
                    _currentLineIndex = startNumber;
                }
                else
                {
                    _currentNodeIndex = startNumber;
                    // 線段編號邏輯：檢查前綴是否改變
                    if (_allLines.Count > 0)
                    {
                        var lastLine = _allLines.Last();
                        string lastPrefix = ExtractPrefix(lastLine.Label);
                        if (lastPrefix == prefix)
                        {
                            int lastNumber = ExtractNumber(lastLine.Label);
                            _currentLineIndex = lastNumber + 1;
                        }
                        else
                        {
                            _currentLineIndex = startNumber;
                        }
                    }
                    _currentSessionNodes.Clear();
                }

                _currentPrefix = prefix;
                _shouldCreateLines = createLines;
                _currentLayerName = $"BLOCK_{prefix}";

                string blockName = "";

                // 使用文檔鎖定進行資料庫操作
                using (DocumentLock docLock = _doc.LockDocument())
                {
                    // 建立或取得圖層
                    _targetLayerId = _blockCreator.CreateOrGetLayer(_currentLayerName);
                    _db.Clayer = _targetLayerId;

                    // 建立節點圖塊定義
                    blockName = $"MyAutoBlock_{DateTime.Now:HHmmss}";
                    _blockCreator.CreateNodeBlockDefinition(blockName, prefix, _currentNodeIndex);
                }

                // 隱藏視窗並開始點選循環（在鎖定外執行）
                _window.Hide();
                ShowPromptMessage();
                ExecuteNodeCreationLoop(blockName);
            }
            catch (System.Exception ex)
            {
                _ed.WriteMessage($"\n錯誤: {ex.Message}\n");
                _window.Show();
            }
        }

        /// <summary>
        /// 顯示提示訊息
        /// </summary>
        private void ShowPromptMessage()
        {
            string message = $"請在圖面上點選插入位置\n" +
                           $"圖塊將建立在圖層：{_currentLayerName}\n" +
                           $"按 ESC 鍵：復原上一個節點\n" +
                           $"輸入 X 或按右鍵選擇「結束」：結束並回到設定視窗";

            Application.ShowAlertDialog(message);
        }

        /// <summary>
        /// 執行節點建立循環
        /// </summary>
        private void ExecuteNodeCreationLoop(string blockName)
        {
            while (true)
            {
                PromptPointOptions ppo = new PromptPointOptions(
                    $"\n請點選節點位置（{_currentLayerName}）"
                );
                ppo.AllowNone = false; // 不允許直接按 Enter

                // 添加關鍵字選項（只保留結束）
                ppo.Keywords.Add("eXit", "X", "結束(X)");
                ppo.AppendKeywordsToMessage = true;

                PromptPointResult ppr = _ed.GetPoint(ppo);

                // 處理使用者輸入
                if (ppr.Status == PromptStatus.Keyword)
                {
                    if (ppr.StringResult == "eXit")
                    {
                        // X - 結束並回到 GUI，清除當前會話節點
                        _ed.WriteMessage("\n已結束節點建立模式。\n");
                        _currentSessionNodes.Clear();
                        _previousNodePoint = null;
                        UpdateNextStartNumber();
                        _window.Show();
                        break;
                    }
                }
                else if (ppr.Status == PromptStatus.Cancel)
                {
                    // ESC - 復原上一個節點
                    if (UndoLastNode())
                    {
                        _ed.WriteMessage("\n已復原上一個節點。\n");
                    }
                    else
                    {
                        _ed.WriteMessage("\n沒有可復原的節點！\n");
                    }
                    continue;
                }
                else if (ppr.Status == PromptStatus.None)
                {
                    // 不應該發生，因為 AllowNone = false
                    continue;
                }
                else if (ppr.Status != PromptStatus.OK)
                {
                    // 其他錯誤 - 退出
                    _ed.WriteMessage("\n已退出節點建立模式。\n");
                    _currentSessionNodes.Clear();
                    _previousNodePoint = null;
                    UpdateNextStartNumber();
                    _window.Show();
                    break;
                }

                Point3d userPoint = ppr.Value;

                // 檢查是否點選到封閉圖形的起點
                if (IsPointDuplicateInCurrentSession(userPoint))
                {
                    // 封閉圖形 - 建立最後一條線段和面域
                    if (_currentSessionNodes.Count > 0 && _previousNodePoint.HasValue)
                    {
                        Point3d firstPoint = _currentSessionNodes[0].Position;
                        if (!DoesLineExistBetweenPoints(_previousNodePoint.Value, firstPoint))
                        {
                            CreateLine(_previousNodePoint.Value, firstPoint);
                        }
                    }

                    CreateAreaRegion(_currentSessionNodes);
                    _currentSessionNodes.Clear();
                    _previousNodePoint = null;
                    continue;
                }

                // 檢查是否點選到已存在的節點
                NodeData existingNode = FindExistingNodeAtPoint(userPoint);
                if (existingNode != null)
                {
                    // 使用已存在的節點
                    _currentSessionNodes.Add(existingNode);

                    // 建立線段
                    if (_shouldCreateLines && _previousNodePoint.HasValue)
                    {
                        if (!DoesLineExistBetweenPoints(_previousNodePoint.Value, userPoint))
                        {
                            CreateLine(_previousNodePoint.Value, userPoint);

                            // 檢查是否形成封閉圖形
                            if (CheckCurrentSessionClosure())
                            {
                                _currentSessionNodes.Clear();
                                _previousNodePoint = null;
                                continue;
                            }
                        }
                    }

                    _previousNodePoint = userPoint;
                }
                else
                {
                    // 建立新節點
                    CreateNode(userPoint, blockName);

                    // 建立線段
                    if (_shouldCreateLines && _previousNodePoint.HasValue)
                    {
                        if (!DoesLineExistBetweenPoints(_previousNodePoint.Value, userPoint))
                        {
                            CreateLine(_previousNodePoint.Value, userPoint);
                        }
                    }

                    _previousNodePoint = userPoint;
                    _currentNodeIndex++;
                }

                // 更新統計
                UpdateStatistics();
            }
        }

        /// <summary>
        /// 建立節點
        /// </summary>
        private void CreateNode(Point3d position, string blockName)
        {
            string label = $"P-{_currentPrefix}-{_currentNodeIndex:D4}";

            using (DocumentLock docLock = _doc.LockDocument())
            {
                // 插入圖塊
                ObjectId blockId = _blockCreator.InsertNodeBlock(position, blockName, _currentLayerName, label);

                // 儲存資料
                NodeData nodeData = new NodeData(label, position, _currentLayerName);
                _allNodes.Add(nodeData);
                _currentSessionNodes.Add(nodeData);
                _nodeObjectIds[label] = blockId;

                _doc.Editor.Regen();
            }
        }

        /// <summary>
        /// 建立線段
        /// </summary>
        private void CreateLine(Point3d startPoint, Point3d endPoint)
        {
            using (DocumentLock docLock = _doc.LockDocument())
            {
                // 檢查當前前綴的線段編號
                int actualLineIndex = GetNextLineIndex();

                string label = $"L-{_currentPrefix}-{actualLineIndex:D4}";
                Point3d midPoint = new Point3d(
                    (startPoint.X + endPoint.X) / 2,
                    (startPoint.Y + endPoint.Y) / 2,
                    0
                );

                // 建立線段圖塊
                string blockName = $"LineBlock_{DateTime.Now:HHmmss}_{actualLineIndex}";
                _blockCreator.CreateLineBlockDefinition(blockName, startPoint, endPoint, midPoint, _currentPrefix, actualLineIndex);

                // 插入圖塊
                ObjectId blockId = _blockCreator.InsertLineBlock(midPoint, blockName, _currentLayerName, label);

                // 儲存資料
                LineData lineData = new LineData(label, startPoint, endPoint, _currentLayerName);
                _allLines.Add(lineData);
                _lineObjectIds[label] = blockId;

                _currentLineIndex = actualLineIndex + 1;
                _doc.Editor.Regen();
            }
        }

        /// <summary>
        /// 建立面域
        /// </summary>
        private void CreateAreaRegion(List<NodeData> nodes)
        {
            if (nodes.Count < 3)
            {
                Application.ShowAlertDialog("至少需要3個點才能建立面域！");
                return;
            }

            using (DocumentLock docLock = _doc.LockDocument())
            {
                int areaIndex = GetNextAreaIndex();
                string label = $"A-{_currentPrefix}-{areaIndex:D4}";

                Point3d centroid = AreaData.CalculateCentroid(nodes);
                double area = AreaData.CalculatePolygonArea(nodes);

                // 建立面域圖塊
                string blockName = $"AreaBlock_{DateTime.Now:HHmmss}_{areaIndex}";
                _blockCreator.CreateAreaBlockDefinition(blockName, nodes, centroid, label);

                // 插入圖塊
                ObjectId blockId = _blockCreator.InsertAreaBlock(centroid, blockName, _currentLayerName, label);

                // 儲存資料
                AreaData areaData = new AreaData(label, centroid, nodes, _currentLayerName, area);
                _allAreas.Add(areaData);
                _areaObjectIds[label] = blockId;

                _doc.Editor.Regen();

                // 更新統計（不顯示彈出視窗）
                _ed.WriteMessage($"\n偵測到封閉圖形！已自動建立面域 {label}\n");
                UpdateStatistics();
            }
        }

        /// <summary>
        /// 復原最後一個節點
        /// </summary>
        private bool UndoLastNode()
        {
            using (DocumentLock docLock = _doc.LockDocument())
            {
                // 優先處理全域節點
                if (_allNodes.Count > 0)
                {
                    var lastNode = _allNodes.Last();
                    string nodePrefix = ExtractPrefix(lastNode.Label);

                    if (nodePrefix == _currentPrefix)
                    {
                        // 復原相關線段
                        UndoRelatedLines(lastNode.Label);

                        // 復原相關面域
                        UndoRelatedAreas(lastNode.Label);

                        // 刪除節點物件
                        if (_nodeObjectIds.ContainsKey(lastNode.Label))
                        {
                            _blockCreator.DeleteEntity(_nodeObjectIds[lastNode.Label]);
                            _nodeObjectIds.Remove(lastNode.Label);
                        }

                        // 從集合中移除
                        _allNodes.Remove(lastNode);
                        _currentSessionNodes.Remove(lastNode);
                        _currentNodeIndex--;

                        // 更新前一個節點位置
                        UpdatePreviousNodePoint();
                        UpdateStatistics();

                        _doc.Editor.Regen();
                        return true;
                    }
                }

                // 處理當前會話節點
                if (_currentSessionNodes.Count > 0)
                {
                    _currentSessionNodes.RemoveAt(_currentSessionNodes.Count - 1);
                    UpdatePreviousNodePoint();
                    return true;
                }

                return false;
            }
        }

        /// <summary>
        /// 復原相關線段
        /// </summary>
        private void UndoRelatedLines(string nodeLabel)
        {
            var linesToRemove = _allLines.Where(line =>
            {
                string startNodeLabel = GetNodeLabelAtPoint(line.StartPoint);
                string endNodeLabel = GetNodeLabelAtPoint(line.EndPoint);
                return startNodeLabel == nodeLabel || endNodeLabel == nodeLabel;
            }).ToList();

            foreach (var line in linesToRemove)
            {
                if (_lineObjectIds.ContainsKey(line.Label))
                {
                    _blockCreator.DeleteEntity(_lineObjectIds[line.Label]);
                    _lineObjectIds.Remove(line.Label);
                }
                _allLines.Remove(line);
            }

            // 更新線段索引
            if (_allLines.Count > 0)
            {
                var lastLine = _allLines.Last();
                int lastNumber = ExtractNumber(lastLine.Label);
                _currentLineIndex = lastNumber + 1;
            }
        }

        /// <summary>
        /// 復原相關面域
        /// </summary>
        private void UndoRelatedAreas(string nodeLabel)
        {
            var areasToRemove = _allAreas.Where(area =>
                area.NodeLabels.Contains(nodeLabel)
            ).ToList();

            foreach (var area in areasToRemove)
            {
                if (_areaObjectIds.ContainsKey(area.Label))
                {
                    _blockCreator.DeleteEntity(_areaObjectIds[area.Label]);
                    _areaObjectIds.Remove(area.Label);
                }

                // 恢復 _currentSessionNodes，以便重新建立面域
                if (area.NodeLabels.Contains(nodeLabel))
                {
                    _currentSessionNodes.Clear();
                    foreach (var label in area.NodeLabels)
                    {
                        var node = _allNodes.FirstOrDefault(n => n.Label == label);
                        if (node != null)
                        {
                            _currentSessionNodes.Add(node);
                        }
                    }
                }

                _allAreas.Remove(area);
            }
        }

        /// <summary>
        /// 更新前一個節點位置
        /// </summary>
        private void UpdatePreviousNodePoint()
        {
            if (_currentSessionNodes.Count > 0)
            {
                _previousNodePoint = _currentSessionNodes.Last().Position;
            }
            else
            {
                // 當 _currentSessionNodes 為空時（例如面域完成後），
                // 不連接到任何節點，避免錯誤連接到已完成的封閉面域
                _previousNodePoint = null;
            }
        }

        /// <summary>
        /// 檢查當前會話是否形成封閉圖形
        /// </summary>
        private bool CheckCurrentSessionClosure()
        {
            if (_currentSessionNodes.Count < 3)
                return false;

            var firstNode = _currentSessionNodes[0];
            var lastNode = _currentSessionNodes.Last();

            if (firstNode.Label == lastNode.Label && _currentSessionNodes.Count >= 3)
            {
                // 移除重複的最後一個節點
                _currentSessionNodes.RemoveAt(_currentSessionNodes.Count - 1);

                // 建立面域
                CreateAreaRegion(_currentSessionNodes);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 檢查點是否在當前會話中重複
        /// </summary>
        private bool IsPointDuplicateInCurrentSession(Point3d point)
        {
            if (_currentSessionNodes.Count < 3)
                return false;

            return _currentSessionNodes.Any(node =>
                BlockCreator.IsPointDuplicate(point, node.Position));
        }

        /// <summary>
        /// 尋找點位上已存在的節點
        /// </summary>
        private NodeData FindExistingNodeAtPoint(Point3d point)
        {
            return _allNodes.FirstOrDefault(node =>
                ExtractPrefix(node.Label) == _currentPrefix &&
                BlockCreator.IsPointDuplicate(point, node.Position));
        }

        /// <summary>
        /// 檢查兩點之間是否已存在線段
        /// </summary>
        private bool DoesLineExistBetweenPoints(Point3d point1, Point3d point2)
        {
            return _allLines.Any(line =>
                ExtractPrefix(line.Label) == _currentPrefix &&
                ((BlockCreator.IsPointDuplicate(point1, line.StartPoint) &&
                  BlockCreator.IsPointDuplicate(point2, line.EndPoint)) ||
                 (BlockCreator.IsPointDuplicate(point1, line.EndPoint) &&
                  BlockCreator.IsPointDuplicate(point2, line.StartPoint))));
        }

        /// <summary>
        /// 取得點位上的節點標籤
        /// </summary>
        private string GetNodeLabelAtPoint(Point3d point)
        {
            var node = _allNodes.FirstOrDefault(n =>
                ExtractPrefix(n.Label) == _currentPrefix &&
                BlockCreator.IsPointDuplicate(point, n.Position));

            return node?.Label ?? "";
        }

        /// <summary>
        /// 取得下一個線段索引
        /// </summary>
        private int GetNextLineIndex()
        {
            var lastLineWithPrefix = _allLines.LastOrDefault(line =>
                ExtractPrefix(line.Label) == _currentPrefix);

            if (lastLineWithPrefix != null)
            {
                int lastNumber = ExtractNumber(lastLineWithPrefix.Label);
                return lastNumber + 1;
            }

            return 1;
        }

        /// <summary>
        /// 取得下一個面域索引
        /// </summary>
        private int GetNextAreaIndex()
        {
            var lastAreaWithPrefix = _allAreas.LastOrDefault(area =>
                ExtractPrefix(area.Label) == _currentPrefix);

            if (lastAreaWithPrefix != null)
            {
                int lastNumber = ExtractNumber(lastAreaWithPrefix.Label);
                return lastNumber + 1;
            }

            return 1;
        }

        /// <summary>
        /// 從標籤中提取前綴
        /// </summary>
        private string ExtractPrefix(string label)
        {
            // 格式: P-ABC-0001 或 L-ABC-0001 或 A-ABC-0001
            var parts = label.Split('-');
            if (parts.Length >= 3)
                return parts[1];
            return "";
        }

        /// <summary>
        /// 從標籤中提取號碼
        /// </summary>
        private int ExtractNumber(string label)
        {
            var parts = label.Split('-');
            if (parts.Length >= 3 && int.TryParse(parts[2], out int number))
                return number;
            return 0;
        }

        /// <summary>
        /// 更新統計資訊
        /// </summary>
        private void UpdateStatistics()
        {
            _window.Dispatcher.Invoke(() =>
            {
                _window.UpdateStatistics(_allNodes.Count, _allLines.Count, _allAreas.Count);
            });
        }

        /// <summary>
        /// 更新下一個起始號碼到 GUI
        /// </summary>
        private void UpdateNextStartNumber()
        {
            _window.Dispatcher.Invoke(() =>
            {
                _window.UpdateStartNumber(_currentNodeIndex);
            });
        }

        /// <summary>
        /// 完成並輸出
        /// </summary>
        public void FinishAndExport()
        {
            try
            {
                // 恢復原始圖層
                if (!string.IsNullOrEmpty(_originalLayerName))
                {
                    using (DocumentLock docLock = _doc.LockDocument())
                    {
                        using (Transaction tr = _db.TransactionManager.StartTransaction())
                        {
                            LayerTable lt = tr.GetObject(_db.LayerTableId, OpenMode.ForRead) as LayerTable;
                            if (lt.Has(_originalLayerName))
                            {
                                _db.Clayer = lt[_originalLayerName];
                            }
                            tr.Commit();
                        }
                    }
                }

                // 輸出到 Excel
                if (_allNodes.Count > 0 || _allLines.Count > 0 || _allAreas.Count > 0)
                {
                    ExcelExporter exporter = new ExcelExporter();
                    string fileName = exporter.ExportToExcel(_allNodes, _allLines, _allAreas, _doc.Name);

                    Application.ShowAlertDialog(
                        $"建立完成！\n" +
                        $"節點數量：{_allNodes.Count} 個\n" +
                        $"線段數量：{_allLines.Count} 條\n" +
                        $"面域數量：{_allAreas.Count} 個\n\n" +
                        $"資料已輸出至Excel：\n{fileName}"
                    );
                }

                // 重置狀態
                ResetSession();
            }
            catch (System.Exception ex)
            {
                _ed.WriteMessage($"\n輸出錯誤: {ex.Message}\n");
            }
        }

        /// <summary>
        /// 重置會話
        /// </summary>
        private void ResetSession()
        {
            _isSessionActive = false;
            _allNodes.Clear();
            _allLines.Clear();
            _allAreas.Clear();
            _currentSessionNodes.Clear();
            _nodeObjectIds.Clear();
            _lineObjectIds.Clear();
            _areaObjectIds.Clear();
            _previousNodePoint = null;

            UpdateStatistics();
        }

        /// <summary>
        /// 清理資源
        /// </summary>
        public void Cleanup()
        {
            // 清理資源
        }
    }
}
