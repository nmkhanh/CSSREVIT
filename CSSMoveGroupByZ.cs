using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Document = Autodesk.Revit.DB.Document;
using Transaction = Autodesk.Revit.DB.Transaction;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  /// <summary>
  /// Move Groups theo ph??ng Z
  /// - Ch?n multiple groups
  /// - Ch?n 1 model line làm reference
  /// - Move group ??n ?i?m projection c?a group point lên model line
  /// - Ch? di chuy?n theo ph??ng Z (gi? nguyên X, Y)
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSMoveGroupByZ : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
    UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
 {
        // Step 1: Select multiple groups
 IList<Reference> groupRefs;
      try
        {
 groupRefs = uidoc.Selection.PickObjects(
     ObjectType.Element,
      new GroupSelectionFilter(),
            "Select groups to move (multiple selection)");
        }
    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
     {
  return Result.Cancelled;
        }

        if (groupRefs == null || groupRefs.Count == 0)
  {
          RevitTaskDialog.Show("Error", "No groups selected.");
   return Result.Failed;
        }

   // Get Group elements
      List<Group> selectedGroups = groupRefs
          .Select(r => doc.GetElement(r))
          .OfType<Group>()
          .ToList();

        if (selectedGroups.Count == 0)
        {
     RevitTaskDialog.Show("Error", "No valid groups selected.");
          return Result.Failed;
        }

    // Step 2: Select reference model line
        Reference lineRef;
        try
     {
          lineRef = uidoc.Selection.PickObject(
            ObjectType.Element,
         new ModelLineSelectionFilter(),
    "Select Model Line as reference for Z elevation");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
       return Result.Cancelled;
        }

        Element lineElement = doc.GetElement(lineRef);
        if (!(lineElement is ModelCurve modelCurve))
        {
          RevitTaskDialog.Show("Error", "Selected element is not a Model Line.");
          return Result.Failed;
        }

        Curve referenceCurve = modelCurve.GeometryCurve;
        if (referenceCurve == null)
        {
          RevitTaskDialog.Show("Error", "Cannot get curve from Model Line.");
          return Result.Failed;
      }

        // Step 3: Move groups
        using (Transaction tx = new Transaction(doc, "Move Groups by Z Projection"))
        {
          tx.Start();

          int successCount = 0;
          List<string> errors = new List<string>();
          Dictionary<string, double> moveDistances = new Dictionary<string, double>();

          foreach (Group group in selectedGroups)
 {
            try
          {
  // Get group location point
              XYZ groupPoint = GetGroupOriginPoint(group);
    if (groupPoint == null)
  {
    errors.Add($"Group {group.Id}: Cannot get origin point");
          continue;
       }

    // Project group point onto reference curve
        XYZ projectedPoint = ProjectPointOntoCurve(groupPoint, referenceCurve);
            if (projectedPoint == null)
   {
         errors.Add($"Group {group.Id}: Cannot project onto curve");
         continue;
              }

         // Calculate Z displacement only (keep X, Y unchanged)
       double deltaZ = projectedPoint.Z - groupPoint.Z;

          // Create translation vector (only Z direction)
     XYZ translation = new XYZ(0, 0, deltaZ);

     // Move group
ElementTransformUtils.MoveElement(doc, group.Id, translation);

       successCount++;
              
   // Track move distance
              string key = group.GroupType.Name;
    if (!moveDistances.ContainsKey(key))
           {
    moveDistances[key] = 0;
              }
        moveDistances[key] = deltaZ; // Store last move distance for this type
            }
            catch (Exception ex)
            {
              errors.Add($"Group {group.Id}: {ex.Message}");
 }
          }

          tx.Commit();

          // Step 4: Show result
      ShowResult(successCount, selectedGroups.Count, moveDistances, errors);
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
 {
        RevitTaskDialog.Show("Error", $"An error occurred:\n{ex.Message}");
        return Result.Failed;
    }
    }

    /// <summary>
    /// Get origin point of group (location point or bounding box center)
    /// </summary>
    private XYZ GetGroupOriginPoint(Group group)
    {
      try
      {
    // Try to get location point first
    LocationPoint locPoint = group.Location as LocationPoint;
 if (locPoint != null)
 {
          return locPoint.Point;
        }

        // Fallback: use bounding box center
        BoundingBoxXYZ bbox = group.get_BoundingBox(null);
        if (bbox != null)
        {
          return (bbox.Min + bbox.Max) / 2.0;
        }

        return null;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error getting group origin: {ex.Message}");
        return null;
  }
    }

    /// <summary>
    /// Project point onto curve (find closest point on curve)
    /// </summary>
  private XYZ ProjectPointOntoCurve(XYZ point, Curve curve)
    {
      try
      {
        // Get closest point on curve
        IntersectionResult result = curve.Project(point);
        
     if (result != null)
        {
        return result.XYZPoint;
      }

        // Fallback: find closest endpoint
        XYZ start = curve.GetEndPoint(0);
        XYZ end = curve.GetEndPoint(1);

        double distToStart = point.DistanceTo(start);
 double distToEnd = point.DistanceTo(end);

return distToStart < distToEnd ? start : end;
      }
      catch (Exception ex)
      {
 System.Diagnostics.Debug.WriteLine($"Error projecting point: {ex.Message}");
        return null;
      }
  }

    /// <summary>
    /// Show result dialog
    /// </summary>
    private void ShowResult(int successCount, int totalGroups, 
          Dictionary<string, double> moveDistances, List<string> errors)
    {
      StringBuilder resultMsg = new StringBuilder();
      
      resultMsg.AppendLine($"? Moved: {successCount} groups");
      resultMsg.AppendLine($"from {totalGroups} selected groups");
      resultMsg.AppendLine($"\nMove Direction: Z only (vertical)");

      if (moveDistances.Count > 0)
      {
        resultMsg.AppendLine($"\nMove Distances:");
        foreach (var kvp in moveDistances.OrderBy(x => x.Key))
   {
          double distanceMm = kvp.Value * 304.8; // Convert feet to mm
    string direction = kvp.Value > 0 ? "?" : "?";
      resultMsg.AppendLine($"  {kvp.Key}: {direction} {Math.Abs(distanceMm):F1} mm");
        }
      }

      if (errors.Count > 0)
      {
        resultMsg.AppendLine($"\n? Failed: {errors.Count}");
        
  foreach (string error in errors.Take(5))
     {
          resultMsg.AppendLine($"  - {error}");
    }
        
        if (errors.Count > 5)
        {
     resultMsg.AppendLine($"  ... and {errors.Count - 5} more");
        }
      }

      RevitTaskDialog.Show("Move Groups Result", resultMsg.ToString());
    }
  }

  /// <summary>
  /// Selection filter for Groups
  /// </summary>


  /// <summary>
  /// Selection filter for Model Lines
  /// </summary>
  public class ModelLineSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem)
    {
// Accept ModelCurve (Model Line)
      if (elem is ModelCurve modelCurve)
      {
        // Ensure it's a line-based curve
        Curve curve = modelCurve.GeometryCurve;
     return curve != null;
      }

return false;
    }

    public bool AllowReference(Reference reference, XYZ position)
    {
      return false;
    }
  }
}
