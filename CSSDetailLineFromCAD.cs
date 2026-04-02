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
using RevitView = Autodesk.Revit.DB.View;

namespace CSSREVIT
{
  /// <summary>
  /// T?o Detail Lines t? CAD Import
  /// - Ch? l?y layers ?ang visible trong view
  /// - Group lines l?i sau khi t?o
  /// - ??t tên group theo tên view
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSDetailLineFromCAD : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;
      RevitView activeView = doc.ActiveView;

      try
    {
        // Step 1: Validate active view
        if (activeView.ViewType == ViewType.ThreeD)
      {
          RevitTaskDialog.Show("Error", "This command cannot be used in 3D views.\n\nPlease switch to a 2D view (Plan, Section, Elevation, Detail).");
  return Result.Failed;
        }

        // Step 2: Select CAD Import
 Reference cadRef;
        try
        {
          cadRef = uidoc.Selection.PickObject(
            ObjectType.Element,
      new CADImportSelectionFilter(),
  "Select CAD Import to convert to Detail Lines");
        }
    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
        return Result.Cancelled;
        }

        Element cadElement = doc.GetElement(cadRef);
   if (!(cadElement is ImportInstance cadImport))
        {
          RevitTaskDialog.Show("Error", "Selected element is not a CAD Import.");
        return Result.Failed;
        }

        // Step 3: Get visible layers
        List<string> visibleLayers = GetVisibleCADLayers(doc, cadImport, activeView);
        
    if (visibleLayers.Count == 0)
      {
          RevitTaskDialog.Show("Warning", "No visible CAD layers found in current view.");
   return Result.Failed;
        }

        // Step 4: Extract curves from visible layers
        List<Curve> curves = ExtractCurvesFromCAD(cadImport, visibleLayers);

        if (curves.Count == 0)
        {
          RevitTaskDialog.Show("Warning", "No curves found in visible CAD layers.");
          return Result.Failed;
    }

   // Step 5: Create Detail Lines and Group
using (Transaction tx = new Transaction(doc, "Create Detail Lines from CAD"))
        {
          tx.Start();

   int successCount = 0;
     List<ElementId> createdLineIds = new List<ElementId>();
    List<string> errors = new List<string>();

   foreach (Curve curve in curves)
   {
       try
       {
              // Create Detail Line
       DetailCurve detailLine = doc.Create.NewDetailCurve(activeView, curve);
    
       if (detailLine != null)
         {
     createdLineIds.Add(detailLine.Id);
        successCount++;
            }
            }
            catch (Exception ex)
            {
     errors.Add($"Curve: {ex.Message}");
   }
    }

          // Step 6: Create Group with view name
     Group lineGroup = null;
   if (createdLineIds.Count > 0)
    {
  try
     {
           lineGroup = doc.Create.NewGroup(createdLineIds);
         
        // Set group name based on view name
    string groupName = GetUniqueGroupName(doc, activeView.Name);
     lineGroup.GroupType.Name = groupName;
    }
            catch (Exception ex)
  {
       errors.Add($"Group creation: {ex.Message}");
            }
          }

          tx.Commit();

          // Step 7: Show result
          ShowResult(successCount, curves.Count, visibleLayers.Count, 
           lineGroup?.GroupType.Name, errors, activeView.Name);
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
    /// Get visible CAD layers in current view
    /// </summary>
    private List<string> GetVisibleCADLayers(Document doc, ImportInstance cadImport, RevitView view)
    {
      List<string> visibleLayers = new List<string>();

      try
      {
      // Get CAD layer table
        GeometryElement geoElement = cadImport.get_Geometry(new Options());
        
        if (geoElement == null)
    return visibleLayers;

   // Get all layer names
        HashSet<string> allLayers = new HashSet<string>();
        
        foreach (GeometryObject geoObj in geoElement)
  {
        if (geoObj is GeometryInstance geoInstance)
          {
            GeometryElement instGeo = geoInstance.GetInstanceGeometry();
            
   foreach (GeometryObject instObj in instGeo)
            {
         GraphicsStyle style = doc.GetElement(instObj.GraphicsStyleId) as GraphicsStyle;
     
     if (style != null && !string.IsNullOrEmpty(style.Name))
           {
   allLayers.Add(style.Name);
 }
  }
     }
        }

        // Check visibility of each layer in current view
        foreach (string layerName in allLayers)
        {
 if (IsLayerVisibleInView(doc, cadImport, layerName, view))
     {
            visibleLayers.Add(layerName);
     }
        }
      }
      catch (Exception ex)
      {
     System.Diagnostics.Debug.WriteLine($"Error getting visible layers: {ex.Message}");
      }

      return visibleLayers;
    }

    /// <summary>
    /// Check if CAD layer is visible in view
    /// </summary>
    private bool IsLayerVisibleInView(Document doc, ImportInstance cadImport, string layerName, RevitView view)
    {
   try
      {
        // Get category for CAD layer
     Category cadCategory = Category.GetCategory(doc, cadImport.Category.Id);
        
    if (cadCategory == null)
  return true; // Default to visible

// Find subcategory for this layer
     foreach (Category subCat in cadCategory.SubCategories)
        {
          if (subCat.Name.EndsWith(layerName, StringComparison.OrdinalIgnoreCase))
          {
     // Check visibility in view
   try
       {
       return view.GetCategoryHidden(subCat.Id) == false;
  }
            catch
      {
              return true; // If can't determine, assume visible
            }
   }
        }

        return true; // Default to visible if not found
   }
      catch
      {
        return true; // Default to visible on error
      }
    }

    /// <summary>
    /// Extract curves from CAD import (only from visible layers)
    /// </summary>
    private List<Curve> ExtractCurvesFromCAD(ImportInstance cadImport, List<string> visibleLayers)
    {
List<Curve> curves = new List<Curve>();

      try
      {
        Options options = new Options
        {
          DetailLevel = ViewDetailLevel.Fine,
          IncludeNonVisibleObjects = false
        };

        GeometryElement geoElement = cadImport.get_Geometry(options);
  
 if (geoElement == null)
       return curves;

        foreach (GeometryObject geoObj in geoElement)
        {
      if (geoObj is GeometryInstance geoInstance)
          {
         GeometryElement instGeo = geoInstance.GetInstanceGeometry();
      ExtractCurvesFromGeometry(instGeo, curves, visibleLayers, cadImport.Document);
          }
        }
      }
  catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error extracting curves: {ex.Message}");
      }

      return curves;
 }

    /// <summary>
    /// Recursively extract curves from geometry
    /// </summary>
    private void ExtractCurvesFromGeometry(GeometryElement geoElement, List<Curve> curves, 
         List<string> visibleLayers, Document doc)
    {
   foreach (GeometryObject geoObj in geoElement)
      {
        // Check if this geometry belongs to a visible layer
    GraphicsStyle style = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;
     
        if (style != null && !IsLayerVisible(style.Name, visibleLayers))
        {
continue; // Skip this geometry - layer not visible
        }

        if (geoObj is Curve curve)
        {
          curves.Add(curve);
        }
   else if (geoObj is PolyLine polyLine)
        {
      // Convert PolyLine to individual Line segments
      IList<XYZ> points = polyLine.GetCoordinates();
      
          for (int i = 0; i < points.Count - 1; i++)
  {
       try
 {
          if (points[i].DistanceTo(points[i + 1]) > 0.001) // Avoid zero-length lines
              {
 curves.Add(Line.CreateBound(points[i], points[i + 1]));
     }
      }
            catch { }
       }
}
        else if (geoObj is GeometryInstance nestedInstance)
        {
          GeometryElement nestedGeo = nestedInstance.GetInstanceGeometry();
 ExtractCurvesFromGeometry(nestedGeo, curves, visibleLayers, doc);
        }
        else if (geoObj is GeometryElement nestedElement)
        {
          ExtractCurvesFromGeometry(nestedElement, curves, visibleLayers, doc);
 }
      }
 }

    /// <summary>
    /// Check if layer is in visible layers list
    /// </summary>
    private bool IsLayerVisible(string styleName, List<string> visibleLayers)
    {
      if (string.IsNullOrEmpty(styleName))
        return false;

      // Check if style name ends with any of the visible layer names
      return visibleLayers.Any(layer => 
        styleName.EndsWith(layer, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Get unique group name based on view name
    /// </summary>
    private string GetUniqueGroupName(Document doc, string viewName)
    {
    string baseName = $"DetailLines_{viewName}";
      string uniqueName = baseName;
      int suffix = 1;

      // Check if name exists
      FilteredElementCollector collector = new FilteredElementCollector(doc);
      var existingGroups = collector
   .OfClass(typeof(GroupType))
    .Cast<GroupType>()
        .Select(gt => gt.Name)
        .ToHashSet();

      while (existingGroups.Contains(uniqueName))
      {
        uniqueName = $"{baseName}_{suffix}";
   suffix++;
      }

      return uniqueName;
    }

  /// <summary>
    /// Show result dialog
    /// </summary>
 private void ShowResult(int successCount, int totalCurves, int visibleLayersCount,
    string groupName, List<string> errors, string viewName)
    {
      StringBuilder resultMsg = new StringBuilder();
      
      resultMsg.AppendLine($"? Created: {successCount} Detail Lines");
      resultMsg.AppendLine($"from {totalCurves} curves in CAD");
      resultMsg.AppendLine($"\nView: {viewName}");
      resultMsg.AppendLine($"Visible Layers: {visibleLayersCount}");
  
      if (!string.IsNullOrEmpty(groupName))
      {
        resultMsg.AppendLine($"\nGroup Name: {groupName}");
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

      RevitTaskDialog.Show("Create Detail Lines Result", resultMsg.ToString());
  }
  }

  /// <summary>
  /// Selection filter for CAD imports
/// </summary>
  public class CADImportSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem)
  {
      return elem is ImportInstance;
    }

    public bool AllowReference(Reference reference, XYZ position)
  {
      return false;
    }
  }
}
