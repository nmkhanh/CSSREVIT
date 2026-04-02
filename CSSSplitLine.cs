using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSSPLITLINE : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      try
      {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // ===== OPTION =====
        TaskDialog td = new TaskDialog("Chọn hướng split");
        td.MainInstruction = "Chọn điểm bắt đầu split ModelLine";
        td.MainContent = "Yes = Start\nNo = End";
        td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No | TaskDialogCommonButtons.Cancel;
        td.DefaultButton = TaskDialogResult.Yes;

        TaskDialogResult result = td.Show();

        if (result == TaskDialogResult.Cancel)
          return Result.Cancelled;

        bool splitFromStart = (result == TaskDialogResult.Yes);

        // ===== PICK MULTI GROUP =====
        IList<Reference> picks = uidoc.Selection.PickObjects(ObjectType.Element, "Pick nhiều Group");

        if (picks == null || picks.Count == 0)
        {
          Trace.WriteLine("Không chọn group nào");
          return Result.Cancelled;
        }

        // ===== READ FILE (1 lần) =====
        string path = @"D:\CSS\CSSREVIT\Rebar.txt";

        if (!File.Exists(path))
        {
          Trace.WriteLine("Không tìm thấy file");
          return Result.Failed;
        }

        var raw = File.ReadAllLines(path)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        if (raw.Count < 2)
        {
          Trace.WriteLine("File sai format");
          return Result.Failed;
        }

        double mmToFt = 1.0 / 304.8;
        double lap = double.Parse(raw[0]) * mmToFt;

        List<double> lengths = raw.Skip(1)
            .Select(x => double.Parse(x) * mmToFt)
            .ToList();

        int n = lengths.Count;

        using (Transaction t = new Transaction(doc, "Split Multi ModelLine"))
        {
          t.Start();

          var a = 0;
          foreach (Reference pick in picks)
          {
            try
            {
              Group group = doc.GetElement(pick) as Group;

              if (group == null)
              {
                Trace.WriteLine("Không phải Group");
                continue;
              }

              // ===== GET MODEL LINE =====
              ModelLine modelLine = group.GetMemberIds()
                  .Select(id => doc.GetElement(id))
                  .OfType<ModelLine>()
                  .FirstOrDefault();

              if (modelLine == null)
              {
                Trace.WriteLine("Group không có ModelLine");
                continue;
              }

              Curve baseCurve = modelLine.GeometryCurve;
              XYZ p0 = baseCurve.GetEndPoint(0);
              XYZ p1 = baseCurve.GetEndPoint(1);

              double totalLength = baseCurve.Length;

              double expected = lengths.Sum() - lap * (n - 1);

              if (Math.Abs(expected - totalLength) > 0.01)
              {
                Trace.WriteLine($"⚠ Sai số chiều dài Group {group.Id}: Input={expected} | Line={totalLength}");
              }

              // ===== DIRECTION =====
              XYZ start = splitFromStart ? p0 : p1;
              XYZ end = splitFromStart ? p1 : p0;

              XYZ dir = (end - start).Normalize();

              SketchPlane sp = modelLine.SketchPlane;
              Plane plane = sp.GetPlane();

              // KHÔNG delete line cũ để tránh lỗi group
              XYZ current = start;

              List<Line> lines = new List<Line>();
              for (int i = 0; i < n; i++)
              {
                try
                {
                  double len = lengths[i];

                  XYZ next = current + dir * len;
                  Line line = Line.CreateBound(current, next);

                  if (!IsCurveInPlane(line, plane))
                  {
                    Trace.WriteLine($"Line lệch plane - Group {group.Id}");
                    continue;
                  }
                  lines.Add(line);
                  

                  if (i < n - 1)
                    current = next - dir * lap;
                  else
                    current = next;
                }
                catch (Exception ex)
                {
                  Trace.WriteLine($"Group {group.Id}: " + ex.Message);
                }
              }

              var models = group.GetMemberIds()
                  .Select(id => doc.GetElement(id))
                  .OfType<ModelLine>().ToList();
              for (int i = 0; i < n; i++)
              {
                models[i].SetSketchPlaneAndCurve(sp, lines[i]);
              }
              a++;
            }
            catch (Exception ex)
            {
              Trace.WriteLine("Group lỗi: " + ex.Message);
            }
          }

          t.Commit();
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
        return Result.Failed;
      }
    }

    private bool IsCurveInPlane(Curve c, Plane p)
    {
      try
      {
        XYZ p0 = c.GetEndPoint(0);
        XYZ p1 = c.GetEndPoint(1);

        double d0 = Math.Abs((p0 - p.Origin).DotProduct(p.Normal));
        double d1 = Math.Abs((p1 - p.Origin).DotProduct(p.Normal));

        return d0 < 1e-6 && d1 < 1e-6;
      }
      catch
      {
        return false;
      }
    }
  }
}