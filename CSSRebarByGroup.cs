using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Document = Autodesk.Revit.DB.Document;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace CSSREVIT
{
  /// <summary>
  /// T?o NHI?U Rebars t? 1 Group ch?a Model Lines/Arcs
  /// Format tên Group: "d@nametype" (ví d?: "16@F2")
  /// - d: ???ng kính rebar (tìm template CSSd, ví d? CSS16)
  /// - nametype: tên type m?i cho rebar
  /// 
  /// FEATURE: T? ??ng phát hi?n và t?o nhi?u rebars t? các c?m curves liên t?c
  /// Ví d?: Group "16@F2" có 12 lines = 4 cây thép F2 (m?i cây 3 lines ch? U)
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSRebarByGroup : IExternalCommand
  {
    private readonly Dictionary<string, RebarBarType> _rebarTypeCache = new Dictionary<string, RebarBarType>();
    private readonly Dictionary<string, List<ElementId>> _rebarGroups = new Dictionary<string, List<ElementId>>();
    private readonly Dictionary<Curve, ModelCurve> _curveToModelCurve = new Dictionary<Curve, ModelCurve>();

    private const double TOLERANCE_CONNECTION = 0.01;  // ~3mm
    private const double TOLERANCE_SORT = 0.0328; // ~10mm

    // ? LOGGING SYSTEM
    private string _logFilePath;
    private List<string> _logMessages = new List<string>();

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      // T?o file log
      string desktopPath = @"D:\CSS";
      _logFilePath = Path.Combine(desktopPath, $"CSSRebarByGroup_Log.txt");
      LogMessage($"=== CSSRebarByGroup Log - {DateTime.Now} ===\n");

      try
      {
        // Select Host
        Reference hostRef;
        try
        {
          hostRef = uidoc.Selection.PickObject(ObjectType.Element, new HostSelectionFilter(),
          "Select host element (Floor, Wall, Column, Foundation, Framing)");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          LogMessage("Operation cancelled by user");
          SaveLog();
          return Result.Cancelled;
        }

        Element host = doc.GetElement(hostRef);
        if (host == null)
        {
          LogMessage("ERROR: Invalid host element");
          SaveLog();
          TaskDialog.Show("Error", "Invalid host element.");
          return Result.Failed;
        }

        LogMessage($"Selected Host: {host.Name} (ID: {host.Id}, Category: {host.Category?.Name})");

        // Select Groups
        IList<Reference> groupRefs;
        try
        {
          groupRefs = uidoc.Selection.PickObjects(ObjectType.Element, new GroupSelectionFilter(),
            "Select groups (format: d@nametype, e.g., 16@F2)");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          LogMessage("Operation cancelled by user");
          SaveLog();
          return Result.Cancelled;
        }

        if (groupRefs == null || groupRefs.Count == 0)
        {
          LogMessage("ERROR: No groups selected");
          SaveLog();
          TaskDialog.Show("Error", "No groups selected.");
          return Result.Failed;
        }

        List<Group> selectedGroups = groupRefs.Select(r => doc.GetElement(r)).OfType<Group>().ToList();

        if (selectedGroups.Count == 0)
        {
          LogMessage("ERROR: No valid groups selected");
          SaveLog();
          TaskDialog.Show("Error", "No valid groups selected.");
          return Result.Failed;
        }

        LogMessage($"\nSelected {selectedGroups.Count} groups:");
        foreach (var g in selectedGroups)
        {
          LogMessage($"  - {g.Name} (ID: {g.Id})");
        }

        // Process Groups
        using (Transaction tx = new Transaction(doc, "Create Rebars from Groups"))
        {
          tx.Start();
          LogMessage("\n=== PROCESSING GROUPS ===");

          int totalRebars = 0;
          List<string> failedGroups = new List<string>();

          foreach (Group group in selectedGroups)
          {
            LogMessage($"\n--- Group: {group.Name} (ID: {group.Id}) ---");
            try
            {
              int rebarsCreated = ProcessGroup(doc, group, host);
              totalRebars += rebarsCreated;
              LogMessage($"? SUCCESS: Created {rebarsCreated} rebars");
            }
            catch (Exception ex)
            {
              string errorMsg = $"{group.Name}: {ex.Message}";
              failedGroups.Add(errorMsg);
              LogMessage($"? FAILED: {errorMsg}");
              LogMessage($"  Stack: {ex.StackTrace}");
            }
          }

          GroupRebarsByName(doc);
          tx.Commit();

          LogMessage($"\n=== SUMMARY ===");
          LogMessage($"Total rebars: {totalRebars}");
          LogMessage($"Failed groups: {failedGroups.Count}");

          SaveLog();
          ShowResult(totalRebars, selectedGroups.Count, failedGroups);
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        LogMessage($"\n=== EXCEPTION ===");
        LogMessage($"Error: {ex.Message}");
        LogMessage($"Stack: {ex.StackTrace}");
        SaveLog();
        TaskDialog.Show("Error", $"An error occurred: {ex.Message}\n\nLog: {_logFilePath}");
        return Result.Failed;
      }
    }

    private void LogMessage(string message)
    {
      _logMessages.Add(message);
      System.Diagnostics.Debug.WriteLine(message);
    }

    private void SaveLog()
    {
      try
      {
        File.WriteAllLines(_logFilePath, _logMessages);
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Failed to save log: {ex.Message}");
      }
    }

    /// <summary>
    /// X? lý 1 Group và t?o NHI?U Rebars t? các c?m curves liên t?c
    /// </summary>
    private int ProcessGroup(Document doc, Group group, Element host)
    {
      string groupName = group.Name;
      LogMessage($"  Parsing: {groupName}");

      // PHÂN TÁCH GROUP NAME
      if (!groupName.Contains("@"))
      {
        LogMessage($"  ERROR: Missing '@'");
        throw new Exception("Group name must follow 'd@nametype' format");
      }

      string[] parts = groupName.Split('@');
      if (parts.Length != 2)
      {
        LogMessage($"  ERROR: Invalid format");
        throw new Exception("Group name must follow 'd@nametype' format");
      }

      string diameter = parts[0].Trim();
      string nameType = parts[1].Trim();
      LogMessage($"  Diameter: {diameter}, Type: {nameType}");

      if (string.IsNullOrEmpty(diameter) || string.IsNullOrEmpty(nameType))
      {
        LogMessage($"  ERROR: Empty values");
        throw new Exception("Group name must follow 'd@nametype' format");
      }

      RebarBarType rebarType = GetOrCreateRebarType(doc, diameter, nameType);
      LogMessage($"  RebarType: {rebarType.Name}");

      // L?Y DANH SÁCH CURVE T? GROUP (v?i ModelCurve mapping)
      _curveToModelCurve.Clear();
      List<Curve> curves = GetCurvesFromGroup(doc, group);
      if (curves.Count == 0)
      {
        LogMessage($"  ERROR: No curves");
        throw new Exception("No model curves found");
      }

      LogMessage($"  Found {curves.Count} curves");

      // ? PHÂN TÁCH thành nhi?u c?m curves liên t?c
      List<List<Curve>> curveGroups = SplitIntoContinuousGroups(curves);
      LogMessage($"  Split into {curveGroups.Count} groups");

      // ? T?O REBAR CHO M?I C?M
      int rebarsCreated = 0;
      for (int i = 0; i < curveGroups.Count; i++)
      {
        LogMessage($"  Group {i + 1}/{curveGroups.Count}: {curveGroups[i].Count} curves");
        try
        {
          List<Curve> sortedCurves = SortCurvesIntoChain(curveGroups[i]);
          ElementId rebarId = CreateRebar(doc, sortedCurves, rebarType, host, nameType);

          if (rebarId != null && rebarId != ElementId.InvalidElementId)
          {
            rebarsCreated++;
            LogMessage($"    ? Created rebar ID={rebarId}");
          }
        }
        catch (Exception ex)
        {
          LogMessage($"    ? Failed: {ex.Message}");
        }
      }

      if (rebarsCreated == 0)
      {
        LogMessage($"  ERROR: No rebars created");
        throw new Exception($"Failed to create any rebars from {curveGroups.Count} groups");
      }

      return rebarsCreated;
    }

    /// <summary>
    /// ? PHÂN TÁCH curves thành các c?m liên t?c riêng bi?t
    /// Ví d?: 
    /// - 12 curves liên t?c (ch? U) ? 4 c?m (m?i c?m 3 curves)
    /// - 10 single lines KHÔNG liên t?c ? 10 c?m (m?i c?m 1 curve)
    /// </summary>
    private List<List<Curve>> SplitIntoContinuousGroups(List<Curve> curves)
    {
      LogMessage($"    SplitIntoContinuousGroups: {curves.Count} curves, tolerance={TOLERANCE_CONNECTION * 304.8:F1}mm");

      List<List<Curve>> groups = new List<List<Curve>>();
      if (curves.Count == 0)
        return groups;

      List<Curve> remaining = new List<Curve>(curves);
      int groupIndex = 0;

      while (remaining.Count > 0)
      {
        groupIndex++;
        // B?t ??u c?m m?i v?i curve ??u tiên
        List<Curve> currentGroup = new List<Curve> { remaining[0] };
        remaining.RemoveAt(0);

        // ? FIX: Ch? tìm curves TH?C S? n?i li?n (< 3mm)
        // KHÔNG gom các curves riêng l? không liên t?c
        bool foundNew = true;
        while (foundNew && remaining.Count > 0)
        {
          foundNew = false;

          // L?y ?i?m ??u và cu?i c?a c?m hi?n t?i
          XYZ groupStart = currentGroup.First().GetEndPoint(0);
          XYZ groupEnd = currentGroup.Last().GetEndPoint(1);

          for (int i = remaining.Count - 1; i >= 0; i--)
          {
            XYZ testStart = remaining[i].GetEndPoint(0);
            XYZ testEnd = remaining[i].GetEndPoint(1);

            // ? KEY FIX: Ch? k?t n?i n?u TH?C S? n?i li?n (< 3mm = TOLERANCE_CONNECTION)
            // Ki?m tra 4 tr??ng h?p k?t n?i
            bool connectsToEnd = groupEnd.DistanceTo(testStart) < TOLERANCE_CONNECTION;
            bool connectsToEndReverse = groupEnd.DistanceTo(testEnd) < TOLERANCE_CONNECTION;
            bool connectsToStart = groupStart.DistanceTo(testEnd) < TOLERANCE_CONNECTION;
            bool connectsToStartReverse = groupStart.DistanceTo(testStart) < TOLERANCE_CONNECTION;

            if (connectsToEnd || connectsToEndReverse || connectsToStart || connectsToStartReverse)
            {
              currentGroup.Add(remaining[i]);
              remaining.RemoveAt(i);
              foundNew = true;

              LogMessage($"      Group {groupIndex}: Added curve (now {currentGroup.Count} curves)");
              break; // Tìm th?y 1 curve n?i li?n, ki?m tra l?i t? ??u
            }
          }
        }

        if (currentGroup.Count > 0)
        {
          groups.Add(currentGroup);
          LogMessage($"      Finished group {groupIndex}: {currentGroup.Count} curve(s)");
        }
      }

      LogMessage($"    Result: {groups.Count} groups total");
      return groups;
    }

    /// <summary>
    /// Ki?m tra curve có n?i v?i b?t k? curve nào trong group không
    /// </summary>
    private bool IsConnectedToGroup(Curve curve, List<Curve> group, double tolerance)
    {
      XYZ cStart = curve.GetEndPoint(0);
      XYZ cEnd = curve.GetEndPoint(1);

      foreach (Curve other in group)
      {
        XYZ oStart = other.GetEndPoint(0);
        XYZ oEnd = other.GetEndPoint(1);

        if (cStart.DistanceTo(oStart) < tolerance || cStart.DistanceTo(oEnd) < tolerance ||
      cEnd.DistanceTo(oStart) < tolerance || cEnd.DistanceTo(oEnd) < tolerance)
          return true;
      }
      return false;
    }

    private List<Curve> SortCurvesIntoChain(List<Curve> curves)
    {
      if (curves.Count <= 1)
        return new List<Curve>(curves);

      List<Curve> sorted = new List<Curve>();
      List<Curve> remaining = new List<Curve>(curves);

      Curve startCurve = FindBestStartCurve(remaining);
      sorted.Add(startCurve);
      remaining.Remove(startCurve);

      XYZ currentEnd = startCurve.GetEndPoint(1);

      while (remaining.Count > 0)
      {
        int bestIndex = -1;
        bool needsReverse = false;
        double minDist = double.MaxValue;

        for (int i = 0; i < remaining.Count; i++)
        {
          double distToStart = currentEnd.DistanceTo(remaining[i].GetEndPoint(0));
          double distToEnd = currentEnd.DistanceTo(remaining[i].GetEndPoint(1));

          if (distToStart < minDist)
          {
            minDist = distToStart;
            bestIndex = i;
            needsReverse = false;
          }
          if (distToEnd < minDist)
          {
            minDist = distToEnd;
            bestIndex = i;
            needsReverse = true;
          }
        }

        if (bestIndex >= 0 && minDist < TOLERANCE_SORT)
        {
          Curve nextCurve = remaining[bestIndex];
          if (needsReverse)
            nextCurve = ReverseCurve(nextCurve);
          sorted.Add(nextCurve);
          currentEnd = nextCurve.GetEndPoint(1);
          remaining.RemoveAt(bestIndex);
        }
        else
        {
          XYZ chainStart = sorted[0].GetEndPoint(0);
          bestIndex = -1;
          minDist = double.MaxValue;

          for (int i = 0; i < remaining.Count; i++)
          {
            double distToStart = chainStart.DistanceTo(remaining[i].GetEndPoint(0));
            double distToEnd = chainStart.DistanceTo(remaining[i].GetEndPoint(1));

            if (distToEnd < minDist)
            {
              minDist = distToEnd;
              bestIndex = i;
              needsReverse = false;
            }
            if (distToStart < minDist)
            {
              minDist = distToStart;
              bestIndex = i;
              needsReverse = true;
            }
          }

          if (bestIndex >= 0 && minDist < TOLERANCE_SORT)
          {
            Curve nextCurve = remaining[bestIndex];
            if (needsReverse)
              nextCurve = ReverseCurve(nextCurve);
            sorted.Insert(0, nextCurve);
            remaining.RemoveAt(bestIndex);
          }
          else
          {
            break;
          }
        }
      }
      return sorted;
    }

    private Curve FindBestStartCurve(List<Curve> curves)
    {
      Curve bestCurve = curves[0];
      int minConnections = int.MaxValue;

      foreach (Curve curve in curves)
      {
        XYZ start = curve.GetEndPoint(0);
        XYZ end = curve.GetEndPoint(1);
        int startConn = 0, endConn = 0;

        foreach (Curve other in curves)
        {
          if (curve == other)
            continue;
          XYZ oStart = other.GetEndPoint(0);
          XYZ oEnd = other.GetEndPoint(1);

          if (start.DistanceTo(oStart) < TOLERANCE_SORT || start.DistanceTo(oEnd) < TOLERANCE_SORT)
            startConn++;
          if (end.DistanceTo(oStart) < TOLERANCE_SORT || end.DistanceTo(oEnd) < TOLERANCE_SORT)
            endConn++;
        }

        int minConn = Math.Min(startConn, endConn);
        if (minConn < minConnections)
        {
          minConnections = minConn;
          bestCurve = curve;
          if (minConn <= 1)
            break;
        }
      }
      return bestCurve;
    }

    private Curve ReverseCurve(Curve curve)
    {
      if (curve is Line line)
        return Line.CreateBound(line.GetEndPoint(1), line.GetEndPoint(0));
      if (curve is Arc arc)
      {
        XYZ mid = arc.Evaluate(0.5, true);
        return Arc.Create(arc.GetEndPoint(1), arc.GetEndPoint(0), mid);
      }
      return curve;
    }

    private ElementId CreateRebar(Document doc, List<Curve> curves, RebarBarType rebarType, Element host, string nameType)
    {
      LogMessage($"    Creating rebar: {curves.Count} curves");

      // ? VALIDATE CURVES
      double totalLength = 0;
      for (int i = 0; i < curves.Count; i++)
      {
        double len = curves[i].Length * 304.8;
        totalLength += len;
        XYZ start = curves[i].GetEndPoint(0);
        XYZ end = curves[i].GetEndPoint(1);
        LogMessage($"      Curve {i + 1}: Length={len:F1}mm, Start=({start.X:F2},{start.Y:F2},{start.Z:F2}), End=({end.X:F2},{end.Y:F2},{end.Z:F2})");
      }
      LogMessage($"      Total length: {totalLength:F1}mm");

      // ? VALIDATE: Minimum length (150mm = ~6 inches)
      if (totalLength < 150)
      {
        LogMessage($"      ? ERROR: Total length {totalLength:F1}mm < 150mm minimum");
        throw new Exception($"Curves too short: {totalLength:F1}mm < 150mm minimum");
      }

      // ? VALIDATE: Maximum length (20m)
      if (totalLength > 20000)
      {
        LogMessage($"      ?? WARNING: Total length {totalLength:F1}mm > 20m (very long rebar)");
      }

      // ? GET HOST INFO
      LogMessage($"      Host: {host.Name} (ID={host.Id}, Category={host.Category?.Name})");
      BoundingBoxXYZ hostBB = host.get_BoundingBox(null);
      if (hostBB != null)
      {
        LogMessage($"      Host BBox: Min=({hostBB.Min.X:F2},{hostBB.Min.Y:F2},{hostBB.Min.Z:F2}), Max=({hostBB.Max.X:F2},{hostBB.Max.Y:F2},{hostBB.Max.Z:F2})");
      }

      // Calculate normal
      XYZ calculatedNormal = CalculateNormal(curves);
      LogMessage($"      Calculated normal: ({calculatedNormal.X:F3}, {calculatedNormal.Y:F3}, {calculatedNormal.Z:F3})");

      List<XYZ> normalsToTry = new List<XYZ>
      {
        calculatedNormal, -calculatedNormal,
        XYZ.BasisZ, -XYZ.BasisZ,
        XYZ.BasisX, -XYZ.BasisX,
        XYZ.BasisY, -XYZ.BasisY
      };

      Rebar rebar = null;
      Exception lastError = null;
      BarTerminationsData barTerminationsData = new BarTerminationsData(doc);

      for (int i = 0; i < 1; i++)
      {
        try
        {
          XYZ normal = normalsToTry[i];

          // ? VALIDATE: Normal must be unit vector
          double normalLen = normal.GetLength();
          if (Math.Abs(normalLen - 1.0) > 0.01)
          {
            LogMessage($"      Attempt {i + 1}: Invalid normal length {normalLen:F3}, normalizing...");
            normal = normal.Normalize();
          }

          // ? VALIDATE: Normal must be perpendicular to first curve
          XYZ curveDir = (curves[0].GetEndPoint(1) - curves[0].GetEndPoint(0)).Normalize();
          double dot = Math.Abs(curveDir.DotProduct(normal));
          LogMessage($"      Attempt {i + 1}: normal=({normal.X:F3},{normal.Y:F3},{normal.Z:F3}), dot={dot:F3}");

          if (dot > 0.1)
          {
            LogMessage($"      ? ? SKIP: Normal NOT perpendicular to curve (dot={dot:F3} > 0.1)");
            continue;
          }

          // ? TRY CREATE REBAR
          XYZ p0 = curves[0].GetEndPoint(0);
          XYZ p1 = curves[0].GetEndPoint(1);
          XYZ mid = p1 + new XYZ(0,0,10);
          Plane plane = Plane.CreateByThreePoints(p0, p1, mid);
          XYZ nor = plane.Normal;
          if(curves.Count > 1)
          {
            XYZ p2 = curves[1].GetEndPoint(1);
            Plane plane2 = Plane.CreateByThreePoints(p1, p2, p0);
            nor = plane2.Normal;
          }
          try
          {
            rebar = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebarType, host, nor, curves,
                      barTerminationsData, true, true);
          }
          catch (Exception createEx)
          {
            // ? LOG DETAILED ERROR FROM REVIT API
            LogMessage($"      ? ? CreateFromCurves FAILED: {createEx.Message}");

            // Check for specific common errors
            if (createEx.Message.Contains("outside"))
            {
              LogMessage($"         ERROR TYPE: Curves outside host boundary");
            }
            else if (createEx.Message.Contains("intersect"))
            {
              LogMessage($"    ERROR TYPE: Curves intersecting/overlapping");
            }
            else if (createEx.Message.Contains("short") || createEx.Message.Contains("length"))
            {
              LogMessage($"    ERROR TYPE: Curves too short");
            }
            else if (createEx.Message.Contains("plane") || createEx.Message.Contains("normal"))
            {
              LogMessage($"         ERROR TYPE: Invalid plane/normal");
            }

            throw; // Re-throw to continue trying other normals
          }

          if (rebar != null)
          {
            LogMessage($"    ? Attempt {i + 1}: SUCCESS with normal ({normal.X:F3},{normal.Y:F3},{normal.Z:F3})");
            LogMessage($"    ? Attempt {i + 1}: SUCCESS with normal ({normal.X:F3},{normal.Y:F3},{normal.Z:F3})");
            break;
          }
        }
        catch (Exception ex)
        {
          lastError = ex;
          LogMessage($"      ? Attempt {i + 1}: {ex.Message}");
        }
      }

      if (rebar == null)
      {
        LogMessage($"");
        LogMessage($"? FAILED after {normalsToTry.Count} attempts");
        LogMessage($"    Curves: {curves.Count}");

        if (curves.Count > 0)
        {
          for (int i = 0; i < curves.Count; i++)
          {
            XYZ start = curves[i].GetEndPoint(0);
            XYZ end = curves[i].GetEndPoint(1);
            LogMessage($"      Curve {i + 1}: Length={curves[i].Length * 304.8:F1}mm");
          }
        }

        string err = $"Could not create rebar from {curves.Count} curve(s)";
        if (lastError != null)
          err += $": {lastError.Message}";
        throw new Exception(err);
      }

      // Set layout
      try
      {
        rebar.GetShapeDrivenAccessor()?.SetLayoutAsFixedNumber(1, 1, true, true, true);
      }
      catch { }

      // Track for grouping
      string baseName = nameType.Contains("-") ? nameType.Substring(0, nameType.IndexOf("-")) : nameType;
      if (!_rebarGroups.ContainsKey(baseName))
        _rebarGroups[baseName] = new List<ElementId>();
      _rebarGroups[baseName].Add(rebar.Id);

      return rebar.Id;
    }

    private XYZ CalculateNormal(List<Curve> curves)
    {
      LogMessage($"   CalculateNormal: {curves.Count} curves");

      // ? CASE 1: Multiple curves - calculate from curve directions
      if (curves.Count >= 2)
      {
        // Try to find normal from consecutive non-parallel curves
        for (int i = 0; i < curves.Count - 1; i++)
        {
          XYZ dir1 = (curves[i].GetEndPoint(1) - curves[i].GetEndPoint(0)).Normalize();
          XYZ dir2 = (curves[i + 1].GetEndPoint(1) - curves[i + 1].GetEndPoint(0)).Normalize();

          double dot = Math.Abs(dir1.DotProduct(dir2));

          // Only use if curves are NOT parallel (dot < 0.99)
          if (dot < 0.99)
          {
            XYZ normal = dir1.CrossProduct(dir2);
            double normalLength = normal.GetLength();

            if (normalLength > 0.001)
            {
              XYZ normalizedNormal = normal.Normalize();
              LogMessage($"        ? Normal from curves {i} and {i + 1}: ({normalizedNormal.X:F3}, {normalizedNormal.Y:F3}, {normalizedNormal.Z:F3})");
              return normalizedNormal;
            }
          }
        }

        LogMessage($"        ? All curves are parallel, using alternative method");
      }

      // ? CASE 2: Single curve OR all curves parallel
      if (curves.Count > 0)
      {
        // Get curve direction
        XYZ curveDir = (curves[0].GetEndPoint(1) - curves[0].GetEndPoint(0)).Normalize();
        LogMessage($"    ? Curve direction: ({curveDir.X:F3}, {curveDir.Y:F3}, {curveDir.Z:F3})");

        // ? TRY: Use SketchPlane normal if available for single curve
        XYZ sketchPlaneNormal = null;
        if (curves.Count == 1 && _curveToModelCurve.TryGetValue(curves[0], out ModelCurve modelCurve))
        {
          try
          {
            SketchPlane sketchPlane = modelCurve.SketchPlane;
            if (sketchPlane != null)
            {
              sketchPlaneNormal = sketchPlane.GetPlane().Normal;
              LogMessage($"  ? SketchPlane normal available: ({sketchPlaneNormal.X:F3}, {sketchPlaneNormal.Y:F3}, {sketchPlaneNormal.Z:F3})");

              // ? VALIDATE: SketchPlane normal must be perpendicular to curve direction
              double dotProduct = Math.Abs(curveDir.DotProduct(sketchPlaneNormal));
              LogMessage($"        ? Dot product (curve · normal): {dotProduct:F3}");

              if (dotProduct < 0.1)  // Nearly perpendicular (dot ? 0 means perpendicular)
              {
                LogMessage($"        ? ? Using SketchPlane normal (perpendicular to curve)");
                return sketchPlaneNormal;
              }
              else
              {
                LogMessage($" ? ? SketchPlane normal NOT perpendicular to curve, calculating alternative");
              }
            }
          }
          catch (Exception ex)
          {
            LogMessage($"        ? Failed to get SketchPlane normal: {ex.Message}");
          }
        }

        // ? IMPROVED LOGIC: Calculate perpendicular normal based on curve orientation

        // Strategy 1: Try cross product with Z axis (works for most horizontal-ish curves)
        XYZ normalZ = curveDir.CrossProduct(XYZ.BasisZ);
        if (normalZ.GetLength() > 0.01)  // Valid normal (not parallel to Z)
        {
          XYZ normalizedZ = normalZ.Normalize();
          LogMessage($"        ? Using curveDir × Z: ({normalizedZ.X:F3}, {normalizedZ.Y:F3}, {normalizedZ.Z:F3})");
          return normalizedZ;
        }

        // Strategy 2: If parallel to Z (vertical curve), use X axis
        LogMessage($"        ? Curve parallel to Z (vertical), using curveDir × X");
        XYZ normalX = curveDir.CrossProduct(XYZ.BasisX);
        if (normalX.GetLength() > 0.01)
        {
          XYZ normalizedX = normalX.Normalize();
          LogMessage($"        ? Using curveDir × X: ({normalizedX.X:F3}, {normalizedX.Y:F3}, {normalizedX.Z:F3})");
          return normalizedX;
        }

        // Strategy 3: If parallel to X, use Y axis
        LogMessage($"        ? Curve parallel to X, using curveDir × Y");
        XYZ normalY = curveDir.CrossProduct(XYZ.BasisY);
        if (normalY.GetLength() > 0.01)
        {
          XYZ normalizedY = normalY.Normalize();
          LogMessage($" ? Using curveDir × Y: ({normalizedY.X:F3}, {normalizedY.Y:F3}, {normalizedY.Z:F3})");
          return normalizedY;
        }
      }

      // ? ABSOLUTE FALLBACK: Use Z basis
      LogMessage($"        ? Absolute fallback: Z basis");
      return XYZ.BasisZ;
    }

    private RebarBarType GetOrCreateRebarType(Document doc, string diameter, string nameType)
    {
      string cacheKey = $"{diameter}@{nameType}";

      if (_rebarTypeCache.TryGetValue(cacheKey, out RebarBarType cachedType))
        return cachedType;

      string templateName = $"CSS{diameter}";
      RebarBarType templateType = FindRebarBarType(doc, templateName);
      if (templateType == null)
        throw new Exception($"Template rebar type '{templateName}' not found");

      RebarBarType existingType = FindRebarBarType(doc, nameType);
      if (existingType != null)
      {
        _rebarTypeCache[cacheKey] = existingType;
        return existingType;
      }

      RebarBarType newType = templateType.Duplicate(nameType) as RebarBarType;
      if (newType == null)
        throw new Exception($"Failed to create rebar type '{nameType}'");

      _rebarTypeCache[cacheKey] = newType;
      return newType;
    }

    private RebarBarType FindRebarBarType(Document doc, string typeName)
    {
      return new FilteredElementCollector(doc)
        .OfClass(typeof(RebarBarType))
   .Cast<RebarBarType>()
        .FirstOrDefault(t => t.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
    }

    private List<Curve> GetCurvesFromGroup(Document doc, Group group)
    {
      List<Curve> curves = new List<Curve>();
      foreach (ElementId id in group.GetMemberIds())
      {
        Element elem = doc.GetElement(id);
        if (elem is ModelCurve modelCurve)
        {
          Curve curve = modelCurve.GeometryCurve;
          if (curve is Line || curve is Arc)
          {
            curves.Add(curve);
            // Store mapping from curve to ModelCurve for later SketchPlane access
            _curveToModelCurve[curve] = modelCurve;
          }
        }
      }
      return curves;
    }

    private void GroupRebarsByName(Document doc)
    {
      foreach (var kvp in _rebarGroups)
      {
        string baseName = kvp.Key;
        List<ElementId> rebarIds = kvp.Value;
        if (rebarIds.Count == 0)
          continue;

        try
        {
          Group group = doc.Create.NewGroup(rebarIds);
          for (int suffix = 0; suffix <= 100; suffix++)
          {
            try
            {
              group.GroupType.Name = suffix == 0 ? baseName : $"{baseName}_{suffix}";
              break;
            }
            catch { }
          }
        }
        catch { }
      }
    }

    private void ShowResult(int totalRebars, int totalGroups, List<string> failedGroups)
    {
      StringBuilder msg = new StringBuilder();
      msg.AppendLine($"? Created: {totalRebars} rebars");
      msg.AppendLine($"   from {totalGroups} groups");
      msg.AppendLine();
      msg.AppendLine($"?? Log: {_logFilePath}");

      if (failedGroups.Count > 0)
      {
        msg.AppendLine();
        msg.AppendLine($"? Failed: {failedGroups.Count} groups");
        foreach (string failed in failedGroups.Take(10))
          msg.AppendLine($"- {failed}");
        if (failedGroups.Count > 10)
          msg.AppendLine($"  ... and {failedGroups.Count - 10} more");
      }

      TaskDialog.Show("Create Rebars Result", msg.ToString());
    }
  }

  public class HostSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem)
    {
      if (elem?.Category == null)
        return false;
      BuiltInCategory cat = (BuiltInCategory)elem.Category.Id.Value;
      return cat == BuiltInCategory.OST_Floors ||
          cat == BuiltInCategory.OST_Walls ||
        cat == BuiltInCategory.OST_StructuralFraming ||
          cat == BuiltInCategory.OST_StructuralColumns ||
     cat == BuiltInCategory.OST_StructuralFoundation ||
          cat == BuiltInCategory.OST_GenericModel;
    }
    public bool AllowReference(Reference reference, XYZ position) => false;
  }

  public class GroupSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem) => elem is Group;
    public bool AllowReference(Reference reference, XYZ position) => false;
  }
}
