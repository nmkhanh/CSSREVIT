using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Document = Autodesk.Revit.DB.Document;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Transaction = Autodesk.Revit.DB.Transaction;
using View = Autodesk.Revit.DB.View;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSCopyGroupToCenterLine : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;
      View activeView = doc.ActiveView;

      try
      {
        // Step 1: Select a group
        Reference groupRef = null;
        try
        {
          groupRef = uidoc.Selection.PickObject(ObjectType.Element, new GroupSelectionFilter(), "Select a group to copy");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          return Result.Cancelled;
        }

        if (groupRef == null)
        {
          TaskDialog.Show("Error", "No group selected.");
          return Result.Failed;
        }

        Element groupElement = doc.GetElement(groupRef);
        if (!(groupElement is Group selectedGroup))
        {
          TaskDialog.Show("Error", "Selected element is not a group.");
          return Result.Failed;
        }

        // Step 2: Find the single CAD import in the active view
        ImportInstance cadImport = FindSingleCADImport(doc, activeView);

        if (cadImport == null)
        {
          TaskDialog.Show("Error", "No CAD import found in the active view, or multiple CAD imports exist.");
          return Result.Failed;
        }

        // Step 3: Get available layers from CAD import
        List<string> layers = GetCADLayers(doc, cadImport);

        if (layers.Count == 0)
        {
          TaskDialog.Show("Error", "No layers found in CAD import.");
          return Result.Failed;
        }

        // Step 4: Let user select a layer
        string selectedLayer = ShowLayerSelectionDialog(layers);

        if (string.IsNullOrEmpty(selectedLayer))
        {
          return Result.Cancelled;
        }

        // Step 5: Get lines from the selected layer
        List<Curve> layerLines = GetLinesFromCADLayer(doc, cadImport, selectedLayer);

        if (layerLines.Count == 0)
        {
          TaskDialog.Show("Error", $"No lines found in layer '{selectedLayer}'.");
          return Result.Failed;
        }

        // Step 6: Get view plane for projection
        Plane viewPlane = GetViewPlane(activeView);

        // Step 7: Copy group to center of each line
        using (Transaction tx = new Transaction(doc, "Copy Group to Center of Lines"))
        {
          tx.Start();

          int copiedCount = 0;
          List<string> errors = new List<string>();

          XYZ groupLocation = GetGroupLocation(selectedGroup);

          foreach (Curve line in layerLines)
          {
            try
            {
              // Get center point of line using start and end points
              XYZ centerPoint = GetCurveCenter(line);

              // Project both points onto view plane to maintain view alignment
              XYZ projectedCenter = ProjectPointOntoPlane(centerPoint, viewPlane);
              XYZ projectedGroupLoc = ProjectPointOntoPlane(groupLocation, viewPlane);

              // Calculate translation vector (only move within view plane)
              XYZ translation = projectedCenter - projectedGroupLoc;

              // Copy the group
              ElementTransformUtils.CopyElement(doc, selectedGroup.Id, translation);
              copiedCount++;
            }
            catch (Exception ex)
            {
              errors.Add($"Line {copiedCount + errors.Count + 1}: {ex.Message}");
            }
          }

          tx.Commit();

          // Show single summary message
          StringBuilder resultMsg = new StringBuilder();
          resultMsg.AppendLine($"✓ Successfully copied: {copiedCount} groups");
          resultMsg.AppendLine($"Total lines in layer: {layerLines.Count}");

          if (errors.Count > 0)
          {
            resultMsg.AppendLine($"\n✗ Failed: {errors.Count}");
            if (errors.Count <= 5)
            {
              resultMsg.AppendLine("\nErrors:");
              foreach (string error in errors)
              {
                resultMsg.AppendLine($"  - {error}");
              }
            }
          }

          TaskDialog.Show("Copy Group Result", resultMsg.ToString());
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
        return Result.Failed;
      }
    }

    private ImportInstance FindSingleCADImport(Document doc, View view)
    {
      // Find all ImportInstance elements in the active view
      FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id);
      List<ImportInstance> cadImports = collector
        .OfClass(typeof(ImportInstance))
        .Cast<ImportInstance>()
        .ToList();

      // Return the single CAD import if exactly one exists
      return cadImports.Count == 1 ? cadImports[0] : null;
    }

    private List<string> GetCADLayers(Document doc, ImportInstance cadImport)
    {
      List<string> layers = new List<string>();

      try
      {
        GeometryElement geoElement = cadImport.get_Geometry(new Options());

        if (geoElement != null)
        {
          foreach (GeometryObject geoObj in geoElement)
          {
            if (geoObj is GeometryInstance geoInstance)
            {
              GeometryElement instanceGeometry = geoInstance.GetInstanceGeometry();

              foreach (GeometryObject instObj in instanceGeometry)
              {
                // Get the category/layer of each geometry object
                if (instObj.GraphicsStyleId != ElementId.InvalidElementId)
                {
                  GraphicsStyle style = doc.GetElement(instObj.GraphicsStyleId) as GraphicsStyle;
                  if (style != null)
                  {
                    // Get the actual layer name from GraphicsStyleCategory
                    Category styleCategory = style.GraphicsStyleCategory;
                    string layerName = styleCategory?.Name ?? style.Name;

                    // Filter out the main import name, we only want actual CAD layers
                    if (!string.IsNullOrEmpty(layerName) && !layers.Contains(layerName))
                    {
                      // Skip if it's just the CAD file name (usually contains .dwg or similar)
                      if (!layerName.Contains(".dwg") && !layerName.Contains(".dxf") &&
                       !layerName.Equals(cadImport.Name, StringComparison.OrdinalIgnoreCase))
                      {
                        layers.Add(layerName);
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error getting CAD layers: {ex.Message}");
      }

      layers.Sort();
      return layers;
    }

    private string ShowLayerSelectionDialog(List<string> layers)
    {
      // For more than 10 layers, use a different approach
      if (layers.Count > 10)
      {
        // Show info and use first 10
        TaskDialog infoDialog = new TaskDialog("Many Layers Found");
        infoDialog.MainInstruction = $"Found {layers.Count} layers in CAD import";
        infoDialog.MainContent = "Showing first 10 layers. The full list is sorted alphabetically.";
        infoDialog.CommonButtons = TaskDialogCommonButtons.Ok;
        infoDialog.Show();
      }

      // Create task dialog to select layer
      TaskDialog td = new TaskDialog("Select CAD Layer");
      td.MainInstruction = "Choose a layer from the CAD import:";
      td.MainContent = $"Found {layers.Count} layer(s). Select the layer containing lines.";
      td.TitleAutoPrefix = false;
      td.AllowCancellation = true;

      // Add layer options as command links (max 10)
      int maxLayers = Math.Min(10, layers.Count);
      for (int i = 0; i < maxLayers; i++)
      {
        td.AddCommandLink((TaskDialogCommandLinkId)(100 + i), layers[i]);
      }

      td.CommonButtons = TaskDialogCommonButtons.Cancel;
      td.DefaultButton = TaskDialogResult.Cancel;

      TaskDialogResult result = td.Show();

      // Get selected layer
      int resultValue = (int)result;
      if (resultValue >= 100 && resultValue < 100 + maxLayers)
      {
        int selectedIndex = resultValue - 100;
        return layers[selectedIndex];
      }

      return null;
    }

    private List<Curve> GetLinesFromCADLayer(Document doc, ImportInstance cadImport, string layerName)
    {
      List<Curve> lines = new List<Curve>();

      try
      {
        GeometryElement geoElement = cadImport.get_Geometry(new Options());

        if (geoElement != null)
        {
          foreach (GeometryObject geoObj in geoElement)
          {
            if (geoObj is GeometryInstance geoInstance)
            {
              GeometryElement instanceGeometry = geoInstance.GetInstanceGeometry();

              foreach (GeometryObject instObj in instanceGeometry)
              {
                // Check if the geometry object is on the selected layer
                if (instObj.GraphicsStyleId != ElementId.InvalidElementId)
                {
                  GraphicsStyle style = doc.GetElement(instObj.GraphicsStyleId) as GraphicsStyle;

                  if (style != null)
                  {
                    // Get layer name from GraphicsStyleCategory
                    Category styleCategory = style.GraphicsStyleCategory;
                    string currentLayerName = styleCategory?.Name ?? style.Name;

                    // IMPORTANT: Only process if this geometry is on the selected layer
                    if (currentLayerName != null && currentLayerName.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                    {
                      // Extract curves (lines) from this geometry
                      if (instObj is Curve curve)
                      {
                        lines.Add(curve);
                      }
                      else if (instObj is PolyLine polyline)
                      {
                        // Convert polyline to individual line segments
                        IList<XYZ> points = polyline.GetCoordinates();
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                          try
                          {
                            Line line = Line.CreateBound(points[i], points[i + 1]);
                            lines.Add(line);
                          }
                          catch { }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error getting lines from layer: {ex.Message}");
      }

      return lines;
    }

    private Plane GetViewPlane(View view)
    {
      // Get the view's origin and direction vectors
      XYZ origin = view.Origin;
      XYZ viewDirection = view.ViewDirection;
      XYZ upDirection = view.UpDirection;
      XYZ rightDirection = viewDirection.CrossProduct(upDirection);

      // Create plane from view direction (normal to the view plane)
      return Plane.CreateByNormalAndOrigin(viewDirection, origin);
    }

    private XYZ ProjectPointOntoPlane(XYZ point, Plane plane)
    {
      // Project point onto plane to maintain alignment with view
      XYZ vectorToPoint = point - plane.Origin;
      double distanceAlongNormal = vectorToPoint.DotProduct(plane.Normal);

      // Subtract the component perpendicular to the plane
      XYZ projectedPoint = point - (distanceAlongNormal * plane.Normal);

      return projectedPoint;
    }

    private XYZ GetGroupLocation(Group group)
    {
      // Get the location of the group
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

    private XYZ GetCurveCenter(Curve curve)
    {
      // Get center point by averaging start and end points
      // This works for all curve types including unbound curves from CAD
      XYZ start = curve.GetEndPoint(0);
      XYZ end = curve.GetEndPoint(1);

      return (start + end) / 2.0;
    }
  }
}
