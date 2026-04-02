using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSExtendModelLine : IExternalCommand
  {
    private string _logFilePath;
    private List<string> _logMessages = new List<string>();

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      // T?o file log
      string desktopPath = @"D:\CSS";
      _logFilePath = Path.Combine(desktopPath, $"CSSExtendModelLine_Log.txt");
      LogMessage($"=== CSSExtendModelLine Log - {DateTime.Now} ===\n");

      try
      {
        // B??c 1: Select group ch?a các line
        Reference groupRef = uidoc.Selection.PickObject(ObjectType.Element, new GroupSelectionFilter(), "B??c 1: Ch?n group ch?a line c?n extend");
        Group selectedGroup = doc.GetElement(groupRef) as Group;

        if (selectedGroup == null)
        {
          TaskDialog.Show("L?i", "Không ph?i group h?p l?");
          return Result.Failed;
        }

        LogMessage($"Selected Group: {selectedGroup.Name} (ID: {selectedGroup.Id})");

        // L?y các model line t? group v?i ElementId
        List<ModelLineData> modelLineDataList = GetModelLinesFromGroup(doc, selectedGroup);
        if (modelLineDataList.Count == 0)
        {
          TaskDialog.Show("L?i", "Không tìm th?y line nào trong group");
          LogMessage("ERROR: Không tìm th?y line nào trong group");
          SaveLog();
          return Result.Failed;
        }

        LogMessage($"Found {modelLineDataList.Count} model lines in group");

        // Phân nhóm các line liên ti?p và l?y ?i?m ??u/cu?i
        List<LineSegment> lineSegments = GetConnectedLineSegments(modelLineDataList);
        LogMessage($"Detected {lineSegments.Count} connected line segments");

        // B??c 2: Select element ?? l?y face
        Reference elementRef = uidoc.Selection.PickObject(ObjectType.Element, $"B??c 2: Ch?n element có face ?? extend ({lineSegments.Count} nhóm line, {modelLineDataList.Count} line t?ng)");
        Element selectedElement = doc.GetElement(elementRef);
        LogMessage($"Selected Element: {selectedElement.Name} (ID: {selectedElement.Id}, Category: {selectedElement.Category?.Name})");

        // L?y t?t c? các face t? element
        List<Face> allFaces = GetFacesFromElement(selectedElement);
        if (allFaces.Count == 0)
        {
          TaskDialog.Show("L?i", "Element không có face nào");
          LogMessage("ERROR: Element không có face nào");
          SaveLog();
          return Result.Failed;
        }

        LogMessage($"Found {allFaces.Count} faces from element");
        for (int i = 0; i < allFaces.Count; i++)
        {
          LogMessage($"  Face {i + 1}: Area = {allFaces[i].Area:F4} sqft ({allFaces[i].Area * 0.092903:F4} m²)");
        }

        // T? ??ng lo?i b? các face có di?n tích b?t th??ng l?n
        List<Face> filteredFaces = FilterOutLargeFaces(allFaces);
        LogMessage($"After filtering: {filteredFaces.Count} faces remain, {allFaces.Count - filteredFaces.Count} large faces removed");

        TaskDialog.Show("Thông tin", $"Element có {allFaces.Count} face\n" +
    $"?ã t? ??ng lo?i {allFaces.Count - filteredFaces.Count} face có di?n tích l?n b?t th??ng\n" +
 $"S? d?ng {filteredFaces.Count} face ?? extend\n\n" +
          $"Log file: {_logFilePath}");

        if (filteredFaces.Count == 0)
        {
          TaskDialog.Show("L?i", "Không còn face nào sau khi l?c");
          LogMessage("ERROR: Không còn face nào sau khi l?c");
          SaveLog();
          return Result.Failed;
        }

        // Th?c hi?n extend
        using (Transaction tx = new Transaction(doc, "Extend Model Lines"))
        {
          tx.Start();

          int extendedSegmentCount = 0;
          int totalExtendedEnds = 0;
          List<string> results = new List<string>();

          LogMessage("\n=== EXTENDING SEGMENTS ===");
          for (int i = 0; i < lineSegments.Count; i++)
          {
            LogMessage($"\nSegment {i + 1}/{lineSegments.Count}: {lineSegments[i].ModelLineDataList.Count} line(s)");
            LogMessage($"  Start: ({lineSegments[i].StartPoint.X:F4}, {lineSegments[i].StartPoint.Y:F4}, {lineSegments[i].StartPoint.Z:F4})");
            LogMessage($"  End: ({lineSegments[i].EndPoint.X:F4}, {lineSegments[i].EndPoint.Y:F4}, {lineSegments[i].EndPoint.Z:F4})");

            results.Add($"Nhóm {i + 1} ({lineSegments[i].ModelLineDataList.Count} line):");
            int extended = ExtendLineSegment(doc, lineSegments[i], filteredFaces, results);
            if (extended > 0)
            {
              extendedSegmentCount++;
              totalExtendedEnds += extended;
            }
          }

          tx.Commit();

          string resultMessage = $"?ã extend {extendedSegmentCount}/{lineSegments.Count} nhóm line\n" +
         $"T?ng {totalExtendedEnds} ??u line ???c extend\n\n" +
                    $"Log file: {_logFilePath}\n\n" +
                    $"Chi ti?t:\n{string.Join("\n", results)}";

          LogMessage($"\n=== SUMMARY ===");
          LogMessage($"Extended {extendedSegmentCount}/{lineSegments.Count} segments");
          LogMessage($"Total {totalExtendedEnds} ends extended");

          SaveLog();
          TaskDialog.Show("Hoàn thành", resultMessage);
        }

        return Result.Succeeded;
      }
      catch (Autodesk.Revit.Exceptions.OperationCanceledException)
      {
        LogMessage("Operation cancelled by user");
        SaveLog();
        return Result.Cancelled;
      }
      catch (Exception ex)
      {
        LogMessage($"EXCEPTION: {ex.Message}");
        LogMessage($"Stack Trace: {ex.StackTrace}");
        SaveLog();
        TaskDialog.Show("L?i", $"?ã x?y ra l?i: {ex.Message}\n\nLog file: {_logFilePath}");
        return Result.Failed;
      }
    }

    private void LogMessage(string message)
    {
      _logMessages.Add(message);
    }

    private void SaveLog()
    {
      try
      {
        File.WriteAllLines(_logFilePath, _logMessages);
      }
      catch
      {
        // Ignore log save errors
      }
    }

    // L?y các model line t? group v?i ElementId
    private List<ModelLineData> GetModelLinesFromGroup(Document doc, Group group)
    {
      List<ModelLineData> modelLines = new List<ModelLineData>();

      foreach (ElementId memberId in group.GetMemberIds())
      {
        Element member = doc.GetElement(memberId);

        // Ki?m tra n?u là ModelCurve
        if (member is ModelCurve modelCurve)
        {
          Curve curve = modelCurve.GeometryCurve;
          if (curve is Line line)
          {
            modelLines.Add(new ModelLineData(memberId, line));
          }
        }
      }

      return modelLines;
    }

    // Phân nhóm các line liên ti?p
    private List<LineSegment> GetConnectedLineSegments(List<ModelLineData> modelLineDataList)
    {
      List<LineSegment> segments = new List<LineSegment>();
      List<ModelLineData> remainingLines = new List<ModelLineData>(modelLineDataList);
      double tolerance = 0.01; // ~3mm tolerance ?? ki?m tra ?i?m trùng

      while (remainingLines.Count > 0)
      {
        List<ModelLineData> connectedGroup = new List<ModelLineData>();
        ModelLineData currentLine = remainingLines[0];
        connectedGroup.Add(currentLine);
        remainingLines.RemoveAt(0);

        bool foundConnection = true;
        while (foundConnection)
        {
          foundConnection = false;
          XYZ startPoint = connectedGroup.First().Line.GetEndPoint(0);
          XYZ endPoint = connectedGroup.Last().Line.GetEndPoint(1);

          for (int i = remainingLines.Count - 1; i >= 0; i--)
          {
            ModelLineData testLineData = remainingLines[i];
            Line testLine = testLineData.Line;
            XYZ testStart = testLine.GetEndPoint(0);
            XYZ testEnd = testLine.GetEndPoint(1);

            // Ki?m tra n?i vào cu?i
            if (endPoint.DistanceTo(testStart) < tolerance)
            {
              connectedGroup.Add(testLineData);
              remainingLines.RemoveAt(i);
              foundConnection = true;
              break;
            }
            else if (endPoint.DistanceTo(testEnd) < tolerance)
            {
              // ??o chi?u line
              Line reversedLine = Line.CreateBound(testEnd, testStart);
              connectedGroup.Add(new ModelLineData(testLineData.ElementId, reversedLine, true));
              remainingLines.RemoveAt(i);
              foundConnection = true;
              break;
            }
            // Ki?m tra n?i vào ??u
            else if (startPoint.DistanceTo(testEnd) < tolerance)
            {
              connectedGroup.Insert(0, testLineData);
              remainingLines.RemoveAt(i);
              foundConnection = true;
              break;
            }
            else if (startPoint.DistanceTo(testStart) < tolerance)
            {
              // ??o chi?u line
              Line reversedLine = Line.CreateBound(testEnd, testStart);
              connectedGroup.Insert(0, new ModelLineData(testLineData.ElementId, reversedLine, true));
              remainingLines.RemoveAt(i);
              foundConnection = true;
              break;
            }
          }
        }

        // T?o segment t? nhóm line liên ti?p
        XYZ segmentStart = connectedGroup.First().Line.GetEndPoint(0);
        XYZ segmentEnd = connectedGroup.Last().Line.GetEndPoint(1);
        segments.Add(new LineSegment(segmentStart, segmentEnd, connectedGroup));
      }

      return segments;
    }

    // L?y t?t c? face t? element
    private List<Face> GetFacesFromElement(Element element)
    {
      List<Face> faces = new List<Face>();
      Options options = new Options();
      options.ComputeReferences = true;
      options.DetailLevel = ViewDetailLevel.Fine;
      options.IncludeNonVisibleObjects = true;

      GeometryElement geomElem = element.get_Geometry(options);

      if (geomElem != null)
      {
        foreach (GeometryObject geomObj in geomElem)
        {
          if (geomObj is Solid solid && solid.Volume > 0)
          {
            foreach (Face face in solid.Faces)
            {
              faces.Add(face);
            }
          }
          else if (geomObj is GeometryInstance geomInst)
          {
            // L?y geometry v?i transform
            GeometryElement instGeom = geomInst.GetInstanceGeometry();
            foreach (GeometryObject instObj in instGeom)
            {
              if (instObj is Solid instSolid && instSolid.Volume > 0)
              {
                foreach (Face face in instSolid.Faces)
                {
                  faces.Add(face);
                }
              }
            }
          }
        }
      }

      return faces;
    }

    // L?c b? các face có di?n tích l?n b?t th??ng
    private List<Face> FilterOutLargeFaces(List<Face> allFaces)
    {
      if (allFaces.Count == 0)
        return allFaces;

      // Tính di?n tích trung bình và ?? l?ch chu?n
      List<double> areas = allFaces.Select(f => f.Area).ToList();
      double avgArea = areas.Average();
      double stdDev = Math.Sqrt(areas.Select(a => Math.Pow(a - avgArea, 2)).Average());

      // Lo?i b? face có di?n tích > trung bình + 2*?? l?ch chu?n
      // Ho?c l?n h?n 3 l?n di?n tích trung bình
      double maxAllowedArea = Math.Min(avgArea + 2 * stdDev, avgArea * 3);

      List<Face> filtered = allFaces.Where(f => f.Area <= maxAllowedArea).ToList();

      // N?u l?c quá nhi?u (>50%), ch? lo?i b? face l?n nh?t
      if (filtered.Count < allFaces.Count / 2)
      {
        var sortedByArea = allFaces.OrderBy(f => f.Area).ToList();
        int keepCount = (int)(allFaces.Count * 0.8); // Gi? 80% face nh? nh?t
        filtered = sortedByArea.Take(keepCount).ToList();
      }

      return filtered;
    }

    // Extend line segment v?i các face
    private int ExtendLineSegment(Document doc, LineSegment segment, List<Face> faces, List<string> results)
    {
      int extendedCount = 0;

      try
      {
        XYZ direction = (segment.EndPoint - segment.StartPoint).Normalize();
        XYZ startPoint = segment.StartPoint;
        XYZ endPoint = segment.EndPoint;

        bool isSingleLine = segment.ModelLineDataList.Count == 1;

        LogMessage($"  Direction: ({direction.X:F4}, {direction.Y:F4}, {direction.Z:F4})");

        // Extend ?i?m ??u (ng??c chi?u)
        XYZ newStartPoint = FindNearestIntersection(startPoint, -direction, faces);

        // Extend ?i?m cu?i (thu?n chi?u)
        XYZ newEndPoint = FindNearestIntersection(endPoint, direction, faces);

        // Debug thông tin
        string debugInfo = $"NewStart: {(newStartPoint != null ? $"Found at {newStartPoint.DistanceTo(startPoint) * 304.8:F1}mm" : "NULL")}, " +
    $"NewEnd: {(newEndPoint != null ? $"Found at {newEndPoint.DistanceTo(endPoint) * 304.8:F1}mm" : "NULL")}";

        LogMessage($"  {debugInfo}");

        // Tr??ng h?p 1: Ch? có 1 line - extend c? 2 ??u
        if (isSingleLine)
        {
          ModelLineData lineData = segment.ModelLineDataList.First();
          ModelCurve modelLine = doc.GetElement(lineData.ElementId) as ModelCurve;

          if (modelLine != null)
          {
            Line originalLine = lineData.Line;
            XYZ originalStart = originalLine.GetEndPoint(0);
            XYZ originalEnd = originalLine.GetEndPoint(1);

            XYZ finalStart = newStartPoint ?? originalStart;
            XYZ finalEnd = newEndPoint ?? originalEnd;

            double startExtendDist = finalStart.DistanceTo(originalStart);
            double endExtendDist = finalEnd.DistanceTo(originalEnd);

            LogMessage($"  Single line extend: Start={startExtendDist * 304.8:F1}mm, End={endExtendDist * 304.8:F1}mm");

            // Extend n?u ít nh?t 1 ??u có giao ?i?m
            if (startExtendDist > 0.001 || endExtendDist > 0.001)
            {
              try
              {
                Line newLine = Line.CreateBound(finalStart, finalEnd);
                modelLine.SetGeometryCurve(newLine, true);

                if (startExtendDist > 0.001)
                  extendedCount++;
                if (endExtendDist > 0.001)
                  extendedCount++;

                results.Add($"  ? 1 line: ??u +{(startExtendDist * 304.8):F1}mm, Cu?i +{(endExtendDist * 304.8):F1}mm");
                LogMessage($"  SUCCESS: Extended single line");
              }
              catch (Exception ex)
              {
                results.Add($"  ? 1 line: L?i - {ex.Message}");
                LogMessage($"  ERROR: {ex.Message}");
              }
            }
            else
            {
              results.Add($"  - 1 line: Không extend ({debugInfo})");
              LogMessage($"  SKIP: No intersection found for single line");
            }
          }
        }
        // Tr??ng h?p 2: Nhi?u line - extend line ??u tiên (?i?m ??u) và line cu?i cùng (?i?m cu?i)
        else
        {
          bool startExtended = false;
          bool endExtended = false;

          // Extend line ??u tiên (ch? ?i?m ??u, gi? nguyên ?i?m cu?i n?i v?i line ti?p theo)
          ModelLineData firstLineData = segment.ModelLineDataList.First();
          ModelCurve firstModelLine = doc.GetElement(firstLineData.ElementId) as ModelCurve;

          if (firstModelLine != null && newStartPoint != null)
          {
            Line firstLine = firstLineData.Line;
            XYZ firstEnd = firstLine.GetEndPoint(1); // Gi? nguyên ?i?m cu?i

            double extendDistance = newStartPoint.DistanceTo(startPoint);
            LogMessage($"  First line extend distance: {extendDistance * 304.8:F1}mm");

            if (extendDistance > 0.001)
            {
              try
              {
                Line newLine = Line.CreateBound(newStartPoint, firstEnd);
                firstModelLine.SetGeometryCurve(newLine, true);
                extendedCount++;
                startExtended = true;
                results.Add($"  ? Line ??u (?i?m ??u): +{(extendDistance * 304.8):F1}mm");
                LogMessage($"  SUCCESS: Extended first line start point");
              }
              catch (Exception ex)
              {
                results.Add($"  ? Line ??u: L?i - {ex.Message}");
                LogMessage($"  ERROR extending first line: {ex.Message}");
              }
            }
          }

          // Extend line cu?i cùng (ch? ?i?m cu?i, gi? nguyên ?i?m ??u n?i v?i line tr??c)
          ModelLineData lastLineData = segment.ModelLineDataList.Last();
          ModelCurve lastModelLine = doc.GetElement(lastLineData.ElementId) as ModelCurve;

          if (lastModelLine != null && newEndPoint != null)
          {
            Line lastLine = lastLineData.Line;
            XYZ lastStart = lastLine.GetEndPoint(0); // Gi? nguyên ?i?m ??u

            double extendDistance = newEndPoint.DistanceTo(endPoint);
            LogMessage($"  Last line extend distance: {extendDistance * 304.8:F1}mm");

            if (extendDistance > 0.001)
            {
              try
              {
                Line newLine = Line.CreateBound(lastStart, newEndPoint);
                lastModelLine.SetGeometryCurve(newLine, true);
                extendedCount++;
                endExtended = true;
                results.Add($"  ? Line cu?i (?i?m cu?i): +{(extendDistance * 304.8):F1}mm");
                LogMessage($"  SUCCESS: Extended last line end point");
              }
              catch (Exception ex)
              {
                results.Add($"  ? Line cu?i: L?i - {ex.Message}");
                LogMessage($"  ERROR extending last line: {ex.Message}");
              }
            }
          }

          if (!startExtended && !endExtended)
          {
            results.Add($"  - Nhóm {segment.ModelLineDataList.Count} line: Không extend ({debugInfo})");
            LogMessage($"  SKIP: No extension for this segment");
          }
        }

        return extendedCount;
      }
      catch (Exception ex)
      {
        results.Add($"  ? L?i: {ex.Message}");
        LogMessage($"  EXCEPTION in ExtendLineSegment: {ex.Message}");
        return 0;
      }
    }

    // Tìm giao ?i?m g?n nh?t v?i các face - FIX: ki?m tra ?úng h??ng
    private XYZ FindNearestIntersection(XYZ fromPoint, XYZ direction, List<Face> faces)
    {
      XYZ nearestPoint = null;
      double minDistance = double.MaxValue;
      int intersectionCount = 0;
      int faceTestedCount = 0;

      LogMessage($"    Finding intersection from ({fromPoint.X:F4}, {fromPoint.Y:F4}, {fromPoint.Z:F4}) in direction ({direction.X:F4}, {direction.Y:F4}, {direction.Z:F4})");

      foreach (Face face in faces)
      {
        faceTestedCount++;
        try
        {
          IntersectionResultArray results = null;

          // Dùng unbounded line ?? tìm t?t c? giao ?i?m
          Line rayUnbounded = Line.CreateUnbound(fromPoint, direction);
          SetComparisonResult result = face.Intersect(rayUnbounded, out results);

          if (result == SetComparisonResult.Overlap && results != null && results.Size > 0)
          {
            foreach (IntersectionResult ir in results)
            {
              XYZ point = ir.XYZPoint;
              double distance = fromPoint.DistanceTo(point);

              // Ki?m tra ?i?m ph?i n?m theo h??ng extend
              XYZ vecToPoint = (point - fromPoint);

              if (distance > 0.001) // Ph?i xa h?n ?i?m g?c
              {
                // Normalize vector tr??c khi tính dotProduct
                XYZ normalizedVec = vecToPoint.Normalize();
                double dotProduct = normalizedVec.DotProduct(direction);

                intersectionCount++;
                LogMessage($"  Face {faceTestedCount} - Intersection {intersectionCount}: distance={distance * 304.8:F1}mm, dotProduct={dotProduct:F3}");

                // CH? ch?p nh?n ?i?m có dotProduct D??NG (cùng h??ng v?i direction)
                // dotProduct > 0: ?i?m n?m theo h??ng extend
                // dotProduct < 0: ?i?m n?m ng??c h??ng ? b? qua
                if (dotProduct > 0.1) // Gi?m threshold xu?ng 0.1 ?? accept nhi?u h??ng h?n
                {
                  if (distance < minDistance)
                  {
                    minDistance = distance;
                    nearestPoint = point;
                    LogMessage($"        -> Selected as nearest (Face {faceTestedCount}, dotProduct={dotProduct:F3})");
                  }
                  else
                  {
                    LogMessage($"      -> Not nearest (current best: {minDistance * 304.8:F1}mm)");
                  }
                }
                else if (dotProduct < 0)
                {
                  LogMessage($"     -> Rejected (opposite direction, dotProduct={dotProduct:F3})");
                }
                else
                {
                  LogMessage($"        -> Rejected (perpendicular, dotProduct={dotProduct:F3})");
                }
              }
              else
              {
                LogMessage($"      Face {faceTestedCount} - Point too close: {distance * 304.8:F1}mm");
              }
            }
          }
        }
        catch (Exception ex)
        {
          LogMessage($"      Exception intersecting face {faceTestedCount}: {ex.Message}");
          continue;
        }
      }

      LogMessage($"    Result: {(nearestPoint != null ? $"Found at {minDistance * 304.8:F1}mm" : "NULL")} (tested {faceTestedCount} faces, {intersectionCount} intersections found)");
      return nearestPoint;
    }
  }

  // Class ?? l?u thông tin ModelLine v?i ElementId
  public class ModelLineData
  {
    public ElementId ElementId
    {
      get; set;
    }
    public Line Line
    {
      get; set;
    }
    public bool IsReversed
    {
      get; set;
    }

    public ModelLineData(ElementId elementId, Line line, bool isReversed = false)
    {
      ElementId = elementId;
      Line = line;
      IsReversed = isReversed;
    }
  }

  // Class ?? l?u thông tin segment line
  public class LineSegment
  {
    public XYZ StartPoint
    {
      get; set;
    }
    public XYZ EndPoint
    {
      get; set;
    }
    public List<ModelLineData> ModelLineDataList
    {
      get; set;
    }

    public LineSegment(XYZ start, XYZ end, List<ModelLineData> modelLineDataList)
    {
      StartPoint = start;
      EndPoint = end;
      ModelLineDataList = modelLineDataList;
    }
  }
}
