using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
  /// Move Groups from their current position to intersection point with selected plane
  /// Step 1: Select plane (face)
  /// Step 2: Select multiple groups
  /// Result: Groups moved VERTICALLY (Z-axis only) until they intersect the plane
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSMoveGroupToPlane : IExternalCommand
  {
    // Logging system
    private string _logFilePath;
    private List<string> _logMessages = new List<string>();

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
    Document doc = uidoc.Document;

      // Create log file
      string desktopPath = @"D:\CSS";
    _logFilePath = Path.Combine(desktopPath, $"CSSMoveGroupToPlane_Log.txt");
      LogMessage($"=== CSSMoveGroupToPlane Log - {DateTime.Now} ===\n");

      try
      {
        // STEP 1: Select Plane (Face)
        Reference planeRef;
 try
        {
          planeRef = uidoc.Selection.PickObject(ObjectType.Face, "Select a planar face to define the target plane");
        }
   catch (Autodesk.Revit.Exceptions.OperationCanceledException)
 {
          LogMessage("Operation cancelled by user at plane selection");
          SaveLog();
          return Result.Cancelled;
  }

if (planeRef == null)
        {
          LogMessage("ERROR: Invalid plane reference");
          SaveLog();
          TaskDialog.Show("Error", "Invalid plane selected.");
          return Result.Failed;
        }

        // Get plane from face
 Element planeElement = doc.GetElement(planeRef);
        GeometryObject geoObj = planeElement.GetGeometryObjectFromReference(planeRef);
        
   Plane targetPlane = null;
      if (geoObj is Face face)
        {
          if (face is PlanarFace planarFace)
          {
       targetPlane = Plane.CreateByNormalAndOrigin(planarFace.FaceNormal, planarFace.Origin);
        LogMessage($"Selected Plane:");
            LogMessage($"  Element: {planeElement.Name} (ID: {planeElement.Id})");
         LogMessage($"  Normal: ({planarFace.FaceNormal.X:F3}, {planarFace.FaceNormal.Y:F3}, {planarFace.FaceNormal.Z:F3})");
        LogMessage($"  Origin: ({planarFace.Origin.X:F3}, {planarFace.Origin.Y:F3}, {planarFace.Origin.Z:F3})");
          }
          else
          {
         LogMessage("ERROR: Selected face is not planar");
 SaveLog();
            TaskDialog.Show("Error", "Please select a planar face.");
  return Result.Failed;
          }
 }
    else
        {
          LogMessage("ERROR: Selected reference is not a face");
          SaveLog();
 TaskDialog.Show("Error", "Please select a face.");
    return Result.Failed;
        }

        // STEP 2: Select Multiple Groups
    IList<Reference> groupRefs;
        try
    {
          groupRefs = uidoc.Selection.PickObjects(ObjectType.Element, new GroupSelectionFilter(),
          "Select groups to move to the plane (vertical movement only)");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
    LogMessage("Operation cancelled by user at group selection");
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
        using (Transaction tx = new Transaction(doc, "Move Groups to Plane (Vertical)"))
        {
       tx.Start();
       LogMessage("\n=== PROCESSING GROUPS ===");
          LogMessage("Movement Mode: VERTICAL ONLY (Z-axis)\n");

        int successCount = 0;
    int failedCount = 0;
       List<string> failedGroups = new List<string>();

  foreach (Group group in selectedGroups)
          {
         LogMessage($"\n--- Group: {group.Name} (ID: {group.Id}) ---");
      try
          {
          XYZ groupLocation = GetGroupLocation(group);
  LogMessage($"  Current location: ({groupLocation.X:F3}, {groupLocation.Y:F3}, {groupLocation.Z:F3})");

        // Calculate Z coordinate on the plane at group's XY position
              double targetZ = CalculateZOnPlane(groupLocation.X, groupLocation.Y, targetPlane);
              
// Create target point (same X, Y but new Z)
              XYZ targetPoint = new XYZ(groupLocation.X, groupLocation.Y, targetZ);
    
              // Calculate move vector (vertical only)
       XYZ moveVector = targetPoint - groupLocation;
         double distance = Math.Abs(moveVector.Z);
       
     LogMessage($"  Target Z: {targetZ:F3} ({targetZ * 304.8:F1}mm)");
     LogMessage($"  Vertical move distance: {distance * 304.8:F1}mm");

            if (distance < 0.0001) // Less than 0.03mm
          {
     LogMessage($"  ?? Group already at target Z, no move needed");
 successCount++;
continue;
    }

  // Move the group vertically
      ElementTransformUtils.MoveElement(doc, group.Id, moveVector);
           
         LogMessage($"  ? SUCCESS: Moved group vertically by {moveVector.Z * 304.8:F1}mm");
       successCount++;
            }
         catch (Exception ex)
         {
    string errorMsg = $"{group.Name}: {ex.Message}";
  failedGroups.Add(errorMsg);
            failedCount++;
   LogMessage($"  ? FAILED: {errorMsg}");
  LogMessage($"    Stack: {ex.StackTrace}");
   }
      }

          tx.Commit();

 LogMessage($"\n=== SUMMARY ===");
    LogMessage($"Success: {successCount}");
          LogMessage($"Failed: {failedCount}");
          LogMessage($"Total: {selectedGroups.Count}");

          SaveLog();
          ShowResult(successCount, failedCount, selectedGroups.Count, failedGroups);
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

    /// <summary>
    /// Get location point of a group
    /// </summary>
    private XYZ GetGroupLocation(Group group)
    {
      LocationPoint locationPoint = group.Location as LocationPoint;
      if (locationPoint != null)
      {
return locationPoint.Point;
      }

      // If no location point, calculate center from bounding box
      BoundingBoxXYZ bbox = group.get_BoundingBox(null);
  if (bbox != null)
      {
     return (bbox.Min + bbox.Max) / 2.0;
  }

   // Default to origin
   return XYZ.Zero;
    }

    /// <summary>
    /// Calculate Z coordinate on plane at given X, Y position
    /// For a plane defined by: A*x + B*y + C*z + D = 0
 /// Where (A, B, C) is the normal and D = -normal.Dot(origin)
    /// Solving for z: z = -(A*x + B*y + D) / C
    /// </summary>
    private double CalculateZOnPlane(double x, double y, Plane plane)
    {
      XYZ normal = plane.Normal;
      XYZ origin = plane.Origin;
    
      // Plane equation: normal.X * (x - origin.X) + normal.Y * (y - origin.Y) + normal.Z * (z - origin.Z) = 0
      // Solving for z:
      // z = origin.Z - (normal.X * (x - origin.X) + normal.Y * (y - origin.Y)) / normal.Z
 
      if (Math.Abs(normal.Z) < 0.0001)
    {
        // Plane is nearly vertical, can't calculate Z intersection
        // Return origin Z as fallback
   LogMessage($"    ?? Warning: Plane is nearly vertical (normal.Z ? 0), using plane origin Z");
     return origin.Z;
      }
      
      double z = origin.Z - (normal.X * (x - origin.X) + normal.Y * (y - origin.Y)) / normal.Z;
  
      return z;
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

    private void ShowResult(int successCount, int failedCount, int totalCount, List<string> failedGroups)
    {
      StringBuilder msg = new StringBuilder();
      msg.AppendLine($"=== Move Groups to Plane Result ===");
    msg.AppendLine($"Movement: VERTICAL ONLY (Z-axis)\n");
      msg.AppendLine($"? Success: {successCount} groups");
      msg.AppendLine($"? Failed: {failedCount} groups");
  msg.AppendLine($"Total: {totalCount} groups");
      msg.AppendLine();
      msg.AppendLine($"?? Log: {_logFilePath}");

      if (failedGroups.Count > 0)
      {
        msg.AppendLine();
        msg.AppendLine($"Failed groups:");
        foreach (string failed in failedGroups.Take(10))
        {
          msg.AppendLine($"  - {failed}");
        }
        if (failedGroups.Count > 10)
  {
    msg.AppendLine($"  ... and {failedGroups.Count - 10} more");
        }
      }

   TaskDialog.Show("Move Groups to Plane", msg.ToString());
    }
  }
}
