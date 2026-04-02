using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSExtendModelLineFinal : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        var gRef = uidoc.Selection.PickObject(ObjectType.Element, "Ch?n Group");
        Group group = doc.GetElement(gRef) as Group;
        if (group == null)
          return Result.Failed;

        var lines = GetLines(doc, group);
        var segments = GroupContinuous(lines);

        var opt = ChooseOption();
        if (opt == TaskDialogResult.Cancel)
          return Result.Cancelled;
        bool isStart = opt == TaskDialogResult.CommandLink1;

        var eRef = uidoc.Selection.PickObject(ObjectType.Element, "Ch?n Element");
        Element elem = doc.GetElement(eRef);

        var faces = FilterFaces(GetFaces(elem));

        using (Transaction tx = new Transaction(doc, "Extend/Trim Line FINAL"))
        {
          tx.Start();

          foreach (var seg in segments)
          {
            var target = isStart ? seg.Lines.First() : seg.Lines.Last();
            ModelCurve mc = doc.GetElement(target.Id) as ModelCurve;
            Line line = mc.GeometryCurve as Line;

            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);

            XYZ fromPoint = isStart ? p0 : p1;
            XYZ keepPoint = isStart ? p1 : p0;

            XYZ hit = FindBestIntersection(line, fromPoint, faces);
            if (hit == null)
              continue;
            XYZ keepPoint_ = keepPoint + new XYZ(0, 0, 10);
            Line line1 = Line.CreateBound(keepPoint, hit);
            Plane plane1 = Plane.CreateByThreePoints(keepPoint_, keepPoint, hit);
            SketchPlane sketchPlane = SketchPlane.Create(doc, plane1);

            //doc.Create.NewModelCurve(line1, sketchPlane);

            //Plane plane = mc.SketchPlane.GetPlane();

            //// ? CÁCH ?ÚNG: Tìm ?i?m G?N NH?T trên line G?C v?i hit
            //// R?i dùng ?i?m ?ó (?ã n?m trên line) ?? t?o line m?i

            //// Extend line g?c v? 2 phía ?? tìm ?i?m g?n nh?t
            //XYZ lineDir = (p1 - p0).Normalize();
            //double lineLength = p0.DistanceTo(p1);

            //// T?o line extend dài ra
            //XYZ extendStart = p0 - lineDir * lineLength * 100; // Extend 100x v? phía p0
            //XYZ extendEnd = p1 + lineDir * lineLength * 100;   // Extend 100x v? phía p1
            //Line extendedLine = Line.CreateBound(extendStart, extendEnd);

            //// Project hit lên extended line ? ?i?m này CH?C CH?C n?m trên line direction
            //IntersectionResult projResult = extendedLine.Project(hit);
            //if (projResult == null)
            //  continue;

            //XYZ hitOnLine = projResult.XYZPoint;

            //// Gi? project 2 ?i?m (??u n?m trên line direction) v? plane
            //XYZ finalKeep = ProjectToPlane(keepPoint, plane);
            //XYZ finalHit = ProjectToPlane(hitOnLine, plane);

            //// T?o line t? 2 ?i?m ?ã project
            //Line newLine = isStart
            //   ? Line.CreateBound(finalHit, finalKeep)
            //      : Line.CreateBound(finalKeep, finalHit);
            mc.SetSketchPlaneAndCurve(sketchPlane, line1);
          }

          tx.Commit();
        }

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
        message = ex.Message;
        return Result.Failed;
      }
    }

    // ================= CORE =================

    private XYZ FindBestIntersection(Line line, XYZ fromPoint, List<Face> faces)
    {
      try
      {
        XYZ s = line.GetEndPoint(0);
        XYZ e = line.GetEndPoint(1);

        XYZ dir = (e - s).Normalize();
        bool isStart = fromPoint.IsAlmostEqualTo(s);
        XYZ extendDir = isStart ? -dir : dir;

        Line bound = line;
        Line unbound = Line.CreateUnbound(fromPoint, extendDir);

        XYZ best = null;
        double minDist = double.MaxValue;

        foreach (var face in faces)
        {
          try
          {
            // TRIM
            IntersectionResultArray arr1;
            if (face.Intersect(bound, out arr1) == SetComparisonResult.Overlap && arr1 != null)
            {
              foreach (IntersectionResult ir in arr1)
              {
                XYZ p = ir.XYZPoint;
                double d = p.DistanceTo(fromPoint);
                if (d < 0.001)
                  continue;
                if (!IsInside(face, p))
                  continue;

                if (d < minDist)
                {
                  minDist = d;
                  best = p;
                }
              }
            }

            // EXTEND
            IntersectionResultArray arr2;
            if (face.Intersect(unbound, out arr2) == SetComparisonResult.Overlap && arr2 != null)
            {
              foreach (IntersectionResult ir in arr2)
              {
                XYZ p = ir.XYZPoint;
                XYZ v = (p - fromPoint);

                double d = v.GetLength();
                if (d < 0.001)
                  continue;
                if (!IsInside(face, p))
                  continue;

                if (v.Normalize().DotProduct(extendDir) < 0.1)
                  continue;

                if (d < minDist)
                {
                  minDist = d;
                  best = p;
                }
              }
            }
          }
          catch { }
        }

        return best;
      }
      catch { return null; }
    }

    private bool IsInside(Face f, XYZ p)
    {
      try
      {
        var r = f.Project(p);
        return r != null && r.Distance < 0.01;
      }
      catch { return false; }
    }

    private XYZ ProjectToPlane(XYZ p, Plane pl)
    {
      double d = (p - pl.Origin).DotProduct(pl.Normal);
      return p - d * pl.Normal;
    }

    // PROJECT ?I?M THEO LINE DIRECTION LÊN PLANE
    private XYZ ProjectAlongLine(XYZ point, XYZ lineDir, Plane plane)
    {
      try
      {
        XYZ n = plane.Normal;
        XYZ o = plane.Origin;
  
        // Line: P = point + t * lineDir
        // Plane: (P - o) · n = 0
        // Solve: t = [(o - point) · n] / (lineDir · n)
      
        double denom = lineDir.DotProduct(n);
     
        // N?u line song song v?i plane ? dùng project vuông góc
        if (Math.Abs(denom) < 0.0001)
          return ProjectToPlane(point, plane);
 
        double t = (o - point).DotProduct(n) / denom;
        return point + t * lineDir;
      }
      catch
      {
        // Fallback: project vuông góc
        return ProjectToPlane(point, plane);
      }
    }

    // ================= DATA =================

    private List<ModelLineData> GetLines(Document doc, Group g)
    {
      var list = new List<ModelLineData>();

      foreach (var id in g.GetMemberIds())
      {
        var mc = doc.GetElement(id) as ModelCurve;
        if (mc?.GeometryCurve is Line l)
          list.Add(new ModelLineData(id, l));
      }

      return list;
    }

    private List<LineSegment> GroupContinuous(List<ModelLineData> input)
    {
      var res = new List<LineSegment>();
      var remain = new List<ModelLineData>(input);
      double tol = 0.01;

      while (remain.Count > 0)
      {
        var group = new List<ModelLineData> { remain[0] };
        remain.RemoveAt(0);

        bool loop = true;
        while (loop)
        {
          loop = false;

          XYZ start = group.First().Line.GetEndPoint(0);
          XYZ end = group.Last().Line.GetEndPoint(1);

          for (int i = remain.Count - 1; i >= 0; i--)
          {
            var l = remain[i].Line;
            XYZ s = l.GetEndPoint(0);
            XYZ e = l.GetEndPoint(1);

            if (end.DistanceTo(s) < tol)
            {
              group.Add(remain[i]);
              remain.RemoveAt(i);
              loop = true;
              break;
            }
            if (end.DistanceTo(e) < tol)
            {
              group.Add(new ModelLineData(remain[i].Id, Line.CreateBound(e, s)));
              remain.RemoveAt(i);
              loop = true;
              break;
            }
            if (start.DistanceTo(e) < tol)
            {
              group.Insert(0, remain[i]);
              remain.RemoveAt(i);
              loop = true;
              break;
            }
            if (start.DistanceTo(s) < tol)
            {
              group.Insert(0, new ModelLineData(remain[i].Id, Line.CreateBound(e, s)));
              remain.RemoveAt(i);
              loop = true;
              break;
            }
          }
        }

        res.Add(new LineSegment(group));
      }

      return res;
    }

    private List<Face> GetFaces(Element e)
    {
      var faces = new List<Face>();
      Options opt = new Options { DetailLevel = ViewDetailLevel.Fine };

      foreach (var obj in e.get_Geometry(opt))
      {
        if (obj is Solid s && s.Volume > 0)
          foreach (Face f in s.Faces)
            faces.Add(f);

        if (obj is GeometryInstance gi)
          foreach (var i in gi.GetInstanceGeometry())
            if (i is Solid s2 && s2.Volume > 0)
              foreach (Face f in s2.Faces)
                faces.Add(f);
      }

      return faces;
    }

    private List<Face> FilterFaces(List<Face> faces)
    {
      if (faces.Count == 0)
        return faces;
      double avg = faces.Average(f => f.Area);
      return faces.Where(f => f.Area < avg * 3).ToList();
    }

    private TaskDialogResult ChooseOption()
    {
      TaskDialog td = new TaskDialog("Ch?n ??u");
      td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "START");
      td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "END");
      td.CommonButtons = TaskDialogCommonButtons.Cancel;
      return td.Show();
    }

    class ModelLineData
    {
      public ElementId Id;
      public Line Line;
      public ModelLineData(ElementId id, Line l)
      {
        Id = id;
        Line = l;
      }
    }

    class LineSegment
    {
      public List<ModelLineData> Lines;
      public LineSegment(List<ModelLineData> l)
      {
        Lines = l;
      }
    }
  }
}