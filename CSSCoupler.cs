using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
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
  [Transaction(TransactionMode.Manual)]
  public class CSSCoupler : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // Step 1: Ask user which end to place coupler
        int endOption = AskUserForEndSelection();
        if (endOption < 0)
        {
          return Result.Cancelled;
        }

        // Step 2: Select MULTIPLE groups containing rebars
        IList<Reference> groupRefs = null;
     try
        {
          groupRefs = uidoc.Selection.PickObjects(ObjectType.Element, new GroupSelectionFilter(), "Select groups containing rebars (multi-select)");
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

        List<Group> selectedGroups = groupRefs
        .Select(r => doc.GetElement(r))
      .OfType<Group>()
      .ToList();

        if (selectedGroups.Count == 0)
        {
      RevitTaskDialog.Show("Error", "No valid groups selected.");
      return Result.Failed;
        }

        // Step 3: Get all rebars from ALL groups
      List<Rebar> rebars = new List<Rebar>();
        foreach (Group g in selectedGroups)
        rebars.AddRange(GetRebarsFromGroup(doc, g));

if (rebars.Count == 0)
      {
   RevitTaskDialog.Show("Error", "No rebars found in the selected groups.");
          return Result.Failed;
        }

        // Step 4: Process rebars and place couplers
        using (Transaction tx = new Transaction(doc, "Place Couplers on Rebars"))
        {
          tx.Start();

          int successCount = 0;
          int failCount = 0;
          List<string> errors = new List<string>();
          Dictionary<string, int> couplersPerDiameter = new Dictionary<string, int>();

          foreach (Rebar rebar in rebars)
          {
            try
            {
              // Get bar diameter from rebar type
              RebarBarType rebarType = doc.GetElement(rebar.GetTypeId()) as RebarBarType;
              if (rebarType == null)
              {
                failCount++;
                errors.Add($"Rebar {rebar.Id}: Could not get rebar type");
                continue;
              }

              // Get Bar Diameter parameter
              Parameter barDiameterParam = rebarType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER);
              if (barDiameterParam == null)
              {
                failCount++;
                errors.Add($"Rebar {rebar.Id}: No Bar Diameter parameter");
                continue;
              }

              double barDiameter = barDiameterParam.AsDouble();
              int diameterMm = (int)Math.Round(barDiameter * 304.8);

              // Find matching coupler type
              ElementType couplerType = FindCouplerType(doc, diameterMm);
              if (couplerType == null)
              {
                failCount++;
                int[] commonDiameters = { 10, 12, 13, 16, 19, 22, 25, 29, 32 };
                int closestDiameter = commonDiameters.OrderBy(d => Math.Abs(d - diameterMm)).First();

                if (closestDiameter != diameterMm)
                {
                  couplerType = FindCouplerType(doc, closestDiameter);
                  if (couplerType != null)
                  {
                    errors.Add($"Rebar {rebar.Id}: Using closest coupler {closestDiameter}mm for actual {diameterMm}mm");
                  }
                  else
                  {
                    errors.Add($"Rebar {rebar.Id}: No coupler type found for diameter {diameterMm}mm (CSS{diameterMm})");
                    continue;
                  }
                }
                else
                {
                  errors.Add($"Rebar {rebar.Id}: No coupler type found for diameter {diameterMm}mm (CSS{diameterMm})");
                  continue;
                }
              }

              // Place coupler at selected end(s)
              bool placed = PlaceCouplerOnRebar(doc, rebar, couplerType, endOption);

              if (placed)
              {
                successCount++;

                string key = $"{diameterMm}mm";
                if (!couplersPerDiameter.ContainsKey(key))
                {
                  couplersPerDiameter[key] = 0;
                }
                couplersPerDiameter[key]++;
              }
              else
              {
                failCount++;
                string endText = endOption == 2 ? "both ends" : $"end {endOption + 1}";
                errors.Add($"Rebar {rebar.Id}: Could not place coupler at {endText}");
              }
            }
            catch (Exception ex)
            {
              failCount++;
              errors.Add($"Rebar {rebar.Id}: {ex.Message}");
            }
          }

          tx.Commit();

          // Show result ONCE
          StringBuilder resultMsg = new StringBuilder();
          resultMsg.AppendLine($"✓ Successfully placed: {successCount} couplers");
          string endDescription = endOption == 0 ? "End 1" : endOption == 1 ? "End 2" : "Both Ends";
          resultMsg.AppendLine($"   at {endDescription}");
          resultMsg.AppendLine($"Total rebars: {rebars.Count}");

          if (couplersPerDiameter.Count > 0)
          {
            resultMsg.AppendLine("\nCouplers by diameter:");
            foreach (var kvp in couplersPerDiameter.OrderBy(x => x.Key))
            {
              resultMsg.AppendLine($"  {kvp.Key}: {kvp.Value} couplers");
            }
          }

          if (failCount > 0)
          {
            resultMsg.AppendLine($"\n✗ Failed: {failCount}");
            if (errors.Count <= 5)
            {
              resultMsg.AppendLine("\nErrors:");
              foreach (string error in errors.Take(5))
              {
                resultMsg.AppendLine($"  - {error}");
              }
            }
            else
            {
              resultMsg.AppendLine($"\nShowing first 5 of {errors.Count} errors:");
              foreach (string error in errors.Take(5))
              {
                resultMsg.AppendLine($"- {error}");
              }
            }
          }

          RevitTaskDialog.Show("Place Couplers Result", resultMsg.ToString());
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        RevitTaskDialog.Show("Error", $"An error occurred: {ex.Message}");
        return Result.Failed;
      }
    }

    private int AskUserForEndSelection()
    {
      Autodesk.Revit.UI.TaskDialog dialog = new Autodesk.Revit.UI.TaskDialog("Select Coupler Placement");
      dialog.MainInstruction = "Choose which end to place couplers:";
      dialog.MainContent = "End 1: Start of rebar (index 0)\nEnd 2: End of rebar (index 1)\nBoth Ends: Place couplers at both ends";

      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "End 1 (Start)", "Place coupler at the start of each rebar");
      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "End 2 (End)", "Place coupler at the end of each rebar");
      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Both Ends", "Place couplers at both ends of each rebar");

      dialog.CommonButtons = TaskDialogCommonButtons.Cancel;
      dialog.DefaultButton = TaskDialogResult.CommandLink1;

      TaskDialogResult result = dialog.Show();

      switch (result)
      {
        case TaskDialogResult.CommandLink1:
          return 0;
        case TaskDialogResult.CommandLink2:
          return 1;
        case TaskDialogResult.CommandLink3:
          return 2;
        default:
          return -1;
      }
    }

    private List<Rebar> GetRebarsFromGroup(Document doc, Group group)
    {
      List<Rebar> rebars = new List<Rebar>();

      try
      {
        ICollection<ElementId> memberIds = group.GetMemberIds();

        foreach (ElementId id in memberIds)
        {
          Element elem = doc.GetElement(id);
          if (elem is Rebar rebar)
          {
            rebars.Add(rebar);
          }
        }
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error getting rebars from group: {ex.Message}");
      }

      return rebars;
    }

    private ElementType FindCouplerType(Document doc, int diameterMm)
    {
      string targetName = $"CSS{diameterMm}";

      FilteredElementCollector collector = new FilteredElementCollector(doc);
      ElementType couplerType = collector
        .OfClass(typeof(ElementType))
            .Cast<ElementType>()
            .FirstOrDefault(ct =>
            {
              if (ct.Category != null && ct.Category.Id.Value == (long)BuiltInCategory.OST_Coupler)
              {
                return ct.FamilyName == "カプラー" && (ct.Name.Contains(targetName, StringComparison.OrdinalIgnoreCase) ||
               ct.Name.EndsWith(targetName, StringComparison.OrdinalIgnoreCase));
              }
              return false;
            });

      return couplerType;
    }

    private bool PlaceCouplerOnRebar(Document doc, Rebar rebar, ElementType couplerType, int endOption)
    {
      try
      {
        IList<Curve> curves = rebar.GetCenterlineCurves(false, false, false, MultiplanarOption.IncludeOnlyPlanarCurves, 0);
        if (curves == null || curves.Count == 0)
        {
          return false;
        }

        bool success = false;

        if (endOption == 0 || endOption == 2)
        {
          bool placed = PlaceCouplerAtEnd(doc, rebar, couplerType, 0);
          success = success || placed;
        }

        if (endOption == 1 || endOption == 2)
        {
          bool placed = PlaceCouplerAtEnd(doc, rebar, couplerType, 1);
          success = success || placed;
        }

        return success;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error placing coupler: {ex.Message}");
        return false;
      }
    }

    private bool PlaceCouplerAtEnd(Document doc, Rebar rebar, ElementType couplerType, int endIndex)
    {
      try
      {
        var reinforcementData = RebarReinforcementData.Create(rebar.Id, endIndex);

        RebarCouplerError error;
        RebarCoupler coupler = RebarCoupler.Create(
          doc,
          couplerType.Id,
       reinforcementData,
          null,
     out error
     );

        if (coupler == null)
        {
          System.Diagnostics.Debug.WriteLine($"Coupler creation error at end {endIndex}: {error}");
          return false;
        }

        return true;
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error placing coupler at end {endIndex}: {ex.Message}");
        return false;
      }
    }
  }
}
