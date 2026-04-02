using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using Document = Autodesk.Revit.DB.Document;
using Transaction = Autodesk.Revit.DB.Transaction;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  #region Data Models

  public class CurveData
  {
    [JsonPropertyName("type")]
    public string Type { get; set; }

    [JsonPropertyName("startPoint")]
    public double[] StartPoint { get; set; }

    [JsonPropertyName("endPoint")]
    public double[] EndPoint { get; set; }

    [JsonPropertyName("midPoint")]
 public double[] MidPoint { get; set; }

    [JsonPropertyName("isBound")]
    public bool IsBound { get; set; }
  }

  public class CurveExportData
  {
[JsonPropertyName("elementName")]
    public string ElementName { get; set; }

    [JsonPropertyName("elementId")]
    public string ElementId { get; set; }

    [JsonPropertyName("exportDate")]
    public string ExportDate { get; set; }

    [JsonPropertyName("curves")]
    public List<CurveData> Curves { get; set; } = new List<CurveData>();

    [JsonPropertyName("totalCurves")]
    public int TotalCurves { get; set; }
  }

  #endregion

  #region Export Command

  [Transaction(TransactionMode.Manual)]
  public class CSSExportCurves : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // Select Generic Model
        Reference selectedRef;
   try
     {
          selectedRef = uidoc.Selection.PickObject(ObjectType.Element,
     new GenericModelSelectionFilter(), "Select Generic Model to export curves");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          return Result.Cancelled;
        }

    Element selectedElement = doc.GetElement(selectedRef);
  if (selectedElement == null)
   {
          RevitTaskDialog.Show("Error", "Invalid element selected.");
          return Result.Failed;
        }

        // Extract curves
        List<Curve> curves = ExtractCurvesFromElement(selectedElement);
    if (curves.Count == 0)
     {
        RevitTaskDialog.Show("Error", "No curves found in selected element.");
          return Result.Failed;
        }

        // Convert to CurveData
        CurveExportData exportData = new CurveExportData
        {
        ElementName = selectedElement.Name,
        ElementId = selectedElement.Id.ToString(),
          ExportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
          TotalCurves = curves.Count
        };

        foreach (Curve curve in curves)
        {
     CurveData curveData = ConvertCurveToCurveData(curve);
          if (curveData != null)
   exportData.Curves.Add(curveData);
        }

  // Show Save File Dialog
        using (SaveFileDialog saveDialog = new SaveFileDialog())
        {
          saveDialog.Title = "Save Curves Export";
          saveDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
          saveDialog.DefaultExt = "json";
          saveDialog.FileName = $"Curves_{selectedElement.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
       saveDialog.InitialDirectory = @"D:\CSS";

   if (saveDialog.ShowDialog() != DialogResult.OK)
    return Result.Cancelled;

          // Export to JSON
          JsonSerializerOptions options = new JsonSerializerOptions
          {
   WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
          };

     string json = JsonSerializer.Serialize(exportData, options);
    File.WriteAllText(saveDialog.FileName, json);

// Show result ONCE
RevitTaskDialog.Show("Export Successful",
   $"? Exported {exportData.TotalCurves} curves\n\n" +
            $"Element: {exportData.ElementName}\n" +
   $"File: {Path.GetFileName(saveDialog.FileName)}\n\n" +
   $"Location: {Path.GetDirectoryName(saveDialog.FileName)}");
   }

        return Result.Succeeded;
}
      catch (Exception ex)
      {
        RevitTaskDialog.Show("Error", $"Export failed:\n{ex.Message}");
 return Result.Failed;
      }
    }

    private List<Curve> ExtractCurvesFromElement(Element element)
    {
      List<Curve> curves = new List<Curve>();
      try
      {
        Options geoOptions = new Options
    {
     DetailLevel = ViewDetailLevel.Fine,
          IncludeNonVisibleObjects = true
      };

        GeometryElement geoElement = element.get_Geometry(geoOptions);
        if (geoElement != null)
        {
 foreach (GeometryObject geoObj in geoElement)
            ExtractCurvesFromGeometry(geoObj, curves);
        }
      }
   catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error extracting curves: {ex.Message}");
      }
return curves;
    }

    private void ExtractCurvesFromGeometry(GeometryObject geoObj, List<Curve> curves)
    {
      if (geoObj is Curve curve)
      {
        curves.Add(curve);
      }
 else if (geoObj is GeometryInstance geoInstance)
 {
        GeometryElement instanceGeo = geoInstance.GetInstanceGeometry();
        if (instanceGeo != null)
        {
          foreach (GeometryObject obj in instanceGeo)
 ExtractCurvesFromGeometry(obj, curves);
        }
      }
      else if (geoObj is GeometryElement geoElement)
      {
        foreach (GeometryObject obj in geoElement)
    ExtractCurvesFromGeometry(obj, curves);
      }
      else if (geoObj is Solid solid)
 {
        foreach (Edge edge in solid.Edges)
        {
        Curve edgeCurve = edge.AsCurve();
          if (edgeCurve != null)
        curves.Add(edgeCurve);
        }
 }
    }

    private CurveData ConvertCurveToCurveData(Curve curve)
    {
      try
   {
        CurveData data = new CurveData { IsBound = curve.IsBound };

        XYZ start = curve.GetEndPoint(0);
        XYZ end = curve.GetEndPoint(1);

        data.StartPoint = new double[] { start.X, start.Y, start.Z };
        data.EndPoint = new double[] { end.X, end.Y, end.Z };

        if (curve is Line)
        {
 data.Type = "Line";
 }
        else if (curve is Arc arc)
  {
     data.Type = "Arc";
          XYZ mid = arc.Evaluate(0.5, true);
          data.MidPoint = new double[] { mid.X, mid.Y, mid.Z };
        }
        else
      {
          data.Type = curve.GetType().Name;
     XYZ mid = curve.Evaluate(0.5, true);
  data.MidPoint = new double[] { mid.X, mid.Y, mid.Z };
        }

        return data;
      }
  catch (Exception ex)
      {
   System.Diagnostics.Debug.WriteLine($"Error converting curve: {ex.Message}");
        return null;
      }
    }
  }

  #endregion

  #region Import Command

  [Transaction(TransactionMode.Manual)]
  public class CSSImportCurves : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // Check if in Family Editor
        if (!doc.IsFamilyDocument)
    {
        RevitTaskDialog.Show("Error", "This command only works in Family Editor.\n\nPlease open a Generic Model Adaptive family.");
          return Result.Failed;
        }

      // Show Open File Dialog
        string filePath;
        using (OpenFileDialog openDialog = new OpenFileDialog())
   {
          openDialog.Title = "Select Curves Export File";
          openDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
          openDialog.DefaultExt = "json";
       openDialog.InitialDirectory = @"D:\CSS";
          openDialog.Multiselect = false;

       if (openDialog.ShowDialog() != DialogResult.OK)
          return Result.Cancelled;

 filePath = openDialog.FileName;
        }

        // Load JSON data
        string json = File.ReadAllText(filePath);
     CurveExportData importData = JsonSerializer.Deserialize<CurveExportData>(json);

        if (importData == null || importData.Curves.Count == 0)
        {
    RevitTaskDialog.Show("Error", "No curves found in JSON file.");
      return Result.Failed;
        }

        // Create Model Curves in Family
     using (Transaction tx = new Transaction(doc, "Import Curves"))
        {
       tx.Start();

  int successCount = 0;
    List<string> errors = new List<string>();

        int curveIndex = 0;
       foreach (CurveData curveData in importData.Curves)
          {
            curveIndex++;
   try
          {
       // Convert to 3D curve
           Curve curve = ConvertCurveDataToCurve(curveData);
      if (curve == null)
          {
            errors.Add($"Curve {curveIndex}: Invalid curve data");
                continue;
       }

           // Create sketch plane that contains this curve
            SketchPlane sketchPlane = CreateSketchPlaneForCurve(doc, curve);
       if (sketchPlane == null)
              {
        errors.Add($"Curve {curveIndex}: Cannot create sketch plane");
    continue;
         }

  // Verify curve is in plane before creating
       if (!IsCurveInPlane(curve, sketchPlane.GetPlane()))
       {
         // Project curve onto plane
     curve = ProjectCurveOntoPlane(curve, sketchPlane.GetPlane());
        if (curve == null)
             {
          errors.Add($"Curve {curveIndex}: Cannot project onto plane");
  continue;
       }
           }

  // Create model curve
              try
    {
    ModelCurve modelCurve = doc.FamilyCreate.NewModelCurve(curve, sketchPlane);
                if (modelCurve != null)
{
   successCount++;
     }
           else
     {
            errors.Add($"Curve {curveIndex}: Failed to create model curve");
      }
              }
        catch (Exception ex)
         {
             errors.Add($"Curve {curveIndex}: {ex.Message}");
   }
   }
            catch (Exception ex)
            {
     errors.Add($"Curve {curveIndex}: {ex.Message}");
     }
          }

          tx.Commit();

          // Show result
          StringBuilder resultMsg = new StringBuilder();
  resultMsg.AppendLine($"? Created: {successCount} model curves");
          resultMsg.AppendLine($"from {importData.TotalCurves} curves in file");
    resultMsg.AppendLine($"\nSource: {importData.ElementName}");
          resultMsg.AppendLine($"Export Date: {importData.ExportDate}");

  if (errors.Count > 0)
          {
            resultMsg.AppendLine($"\n? Failed: {errors.Count}");
  foreach (string error in errors.Take(10))
           resultMsg.AppendLine($"  - {error}");
            if (errors.Count > 10)
              resultMsg.AppendLine($"  ... and {errors.Count - 10} more");
          }

          RevitTaskDialog.Show("Import Result", resultMsg.ToString());
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
 {
        RevitTaskDialog.Show("Error", $"Import failed:\n{ex.Message}");
        return Result.Failed;
      }
    }

    /// <summary>
    /// Check if curve lies in the plane
    /// </summary>
    private bool IsCurveInPlane(Curve curve, Plane plane)
    {
      try
      {
        const double tolerance = 0.001; // ~0.3mm

        XYZ start = curve.GetEndPoint(0);
        XYZ end = curve.GetEndPoint(1);

        // Check if both endpoints are in plane
        double distStart = Math.Abs(plane.Normal.DotProduct(start - plane.Origin));
        double distEnd = Math.Abs(plane.Normal.DotProduct(end - plane.Origin));

        if (distStart > tolerance || distEnd > tolerance)
          return false;

        // For arcs, also check mid point
        if (curve is Arc arc)
        {
  XYZ mid = arc.Evaluate(0.5, true);
          double distMid = Math.Abs(plane.Normal.DotProduct(mid - plane.Origin));
          if (distMid > tolerance)
       return false;
        }

      return true;
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Create sketch plane that contains the curve
 /// </summary>
    private SketchPlane CreateSketchPlaneForCurve(Document doc, Curve curve)
  {
      try
      {
        XYZ start = curve.GetEndPoint(0);
        XYZ end = curve.GetEndPoint(1);
    XYZ direction = (end - start).Normalize();

        XYZ normal;

        if (curve is Arc arc)
        {
   // For arcs, use arc's natural normal
       try
          {
    normal = arc.Normal;
  }
        catch
     {
     // Calculate normal from 3 points
  XYZ mid = arc.Evaluate(0.5, true);
  XYZ v1 = (mid - start).Normalize();
  XYZ v2 = (end - start).Normalize();
            normal = v1.CrossProduct(v2);

            if (normal.GetLength() < 0.001)
         normal = XYZ.BasisZ;
            else
  normal = normal.Normalize();
          }
        }
        else
        {
      // For lines, find best perpendicular direction
          if (Math.Abs(direction.DotProduct(XYZ.BasisZ)) < 0.9)
          {
      // Not parallel to Z - use Z as normal
            normal = XYZ.BasisZ;
    }
       else
  {
            // Nearly vertical - use X as normal
    normal = XYZ.BasisX;
}
        }

        // Create plane at curve start point
        Plane plane = Plane.CreateByNormalAndOrigin(normal, start);
        return SketchPlane.Create(doc, plane);
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error creating sketch plane: {ex.Message}");

        // Fallback
        try
        {
          Plane defaultPlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero);
          return SketchPlane.Create(doc, defaultPlane);
        }
        catch
{
          return null;
        }
      }
    }

 /// <summary>
    /// Project curve onto plane
    /// </summary>
private Curve ProjectCurveOntoPlane(Curve curve, Plane plane)
    {
      try
      {
     XYZ start = curve.GetEndPoint(0);
   XYZ end = curve.GetEndPoint(1);

        XYZ projStart = ProjectPointOntoPlane(start, plane);
        XYZ projEnd = ProjectPointOntoPlane(end, plane);

   // Check distance
        if (projStart.DistanceTo(projEnd) < 0.001)
      return null;

        if (curve is Line)
        {
          return Line.CreateBound(projStart, projEnd);
        }
        else if (curve is Arc arc)
        {
     XYZ mid = arc.Evaluate(0.5, true);
        XYZ projMid = ProjectPointOntoPlane(mid, plane);

          try
   {
  return Arc.Create(projStart, projEnd, projMid);
      }
          catch
          {
      // Arc creation failed, return line
        return Line.CreateBound(projStart, projEnd);
     }
        }
        else
        {
          return Line.CreateBound(projStart, projEnd);
    }
      }
      catch
  {
        return null;
      }
    }

    /// <summary>
    /// Project point onto plane
 /// </summary>
    private XYZ ProjectPointOntoPlane(XYZ point, Plane plane)
    {
      XYZ normal = plane.Normal;
      XYZ origin = plane.Origin;
      XYZ v = point - origin;
      double distance = v.DotProduct(normal);
      return point - (distance * normal);
    }

    /// <summary>
    /// Convert JSON data to Curve
    /// </summary>
    private Curve ConvertCurveDataToCurve(CurveData data)
    {
      try
      {
        // Validate
        if (data == null)
        {
          System.Diagnostics.Debug.WriteLine("CurveData is null");
  return null;
        }

if (data.StartPoint == null || data.StartPoint.Length != 3)
    {
   System.Diagnostics.Debug.WriteLine($"Invalid StartPoint");
 return null;
        }

 if (data.EndPoint == null || data.EndPoint.Length != 3)
        {
     System.Diagnostics.Debug.WriteLine($"Invalid EndPoint");
          return null;
        }

 // Create points
   XYZ start = new XYZ(data.StartPoint[0], data.StartPoint[1], data.StartPoint[2]);
XYZ end = new XYZ(data.EndPoint[0], data.EndPoint[1], data.EndPoint[2]);

        // Check distance
        double distance = start.DistanceTo(end);
        if (distance < 0.0001) // < 0.03mm
        {
          System.Diagnostics.Debug.WriteLine($"Points too close: {distance * 304.8:F6} mm");
          return null;
        }

        // Create curve
        if (data.Type == "Line")
        {
 return Line.CreateBound(start, end);
    }
        else if (data.Type == "Arc")
        {
          if (data.MidPoint == null || data.MidPoint.Length != 3)
          {
            System.Diagnostics.Debug.WriteLine("Arc missing MidPoint - creating Line");
        return Line.CreateBound(start, end);
          }

          XYZ mid = new XYZ(data.MidPoint[0], data.MidPoint[1], data.MidPoint[2]);

          // Check collinearity
        XYZ v1 = (mid - start).Normalize();
          XYZ v2 = (end - start).Normalize();
    double dot = Math.Abs(v1.DotProduct(v2));

      if (dot > 0.999)
          {
     System.Diagnostics.Debug.WriteLine("Arc points collinear - creating Line");
            return Line.CreateBound(start, end);
          }

          try
  {
 return Arc.Create(start, end, mid);
          }
          catch (Exception ex)
   {
            System.Diagnostics.Debug.WriteLine($"Arc creation failed: {ex.Message} - creating Line");
        return Line.CreateBound(start, end);
   }
        }
        else
        {
          System.Diagnostics.Debug.WriteLine($"Unknown curve type: {data.Type} - creating Line");
 return Line.CreateBound(start, end);
        }
      }
  catch (Exception ex)
      {
      System.Diagnostics.Debug.WriteLine($"Error converting curve: {ex.Message}");
return null;
      }
    }
  }

  #endregion

  #region Selection Filters

  public class GenericModelSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem)
    {
      if (elem?.Category == null)
  return false;

      return elem.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel;
 }

    public bool AllowReference(Reference reference, XYZ position)
    {
      return false;
    }
  }

  #endregion
}
