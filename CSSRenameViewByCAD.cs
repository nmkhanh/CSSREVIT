using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Transactions;
using System.Xml.Linq;
using Document = Autodesk.Revit.DB.Document;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Transaction = Autodesk.Revit.DB.Transaction;
using View = Autodesk.Revit.DB.View;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSRenameViewByCAD : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      Selection sel = uidoc.Selection;
      foreach (ElementId id in sel.GetElementIds())
      {
        Element elem = doc.GetElement(id);
        if (elem is View view)
        {
          change(doc, view);
        }
      }

      return Result.Succeeded;
    }

    public void change(Document doc, View activeView)
    {
      try
      {
        if (activeView == null)
        {
          TaskDialog.Show("Error", "No active view found.");
        }

        // Find all ImportInstance elements in the active view
        FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
        ImportInstance firstCadImport = collector
          .OfClass(typeof(ImportInstance))
          .Cast<ImportInstance>()
          .FirstOrDefault();

        if (firstCadImport == null)
        {
          TaskDialog.Show("Error", "No CAD import found in the active view.");
        }

        // Get the CAD link type to retrieve the file name
        CADLinkType cadLinkType = doc.GetElement(firstCadImport.GetTypeId()) as CADLinkType;

        if (cadLinkType == null)
        {
          TaskDialog.Show("Error", "Could not retrieve CAD link type.");
        }

        // Get the file name without extension
        string cadFileName = Path.GetFileNameWithoutExtension(cadLinkType.Name);

        // Rename the view
        using (Transaction tx = new Transaction(doc, "Rename View by CAD Import"))
        {
          tx.Start();

          try
          {
            activeView.Name = cadFileName;
            tx.Commit();
          }
          catch (Exception ex)
          {
            tx.RollBack();
            TaskDialog.Show("Error", $"Failed to rename view: {ex.Message}");
          }
        }
      }
      catch (Exception ex)
      {
        TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
      }
    }
  }
}