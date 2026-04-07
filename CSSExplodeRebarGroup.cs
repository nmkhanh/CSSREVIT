using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Document = Autodesk.Revit.DB.Document;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace CSSREVIT
{
  /// <summary>
  /// Tách các cây thép có Quantity > 1 (NumberOfBarPositions > 1) thành từng cây riêng lẻ,
  /// sau đó group lại theo RebarBarType và pin nhóm.
  /// Cây thép Quantity = 1 được giữ nguyên nhưng vẫn được gộp vào group theo type.
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSExplodeRebarGroup : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // ===== PICK MULTI REBAR =====
        IList<Reference> picks;
        try
        {
          picks = uidoc.Selection.PickObjects(
            ObjectType.Element,
            new RebarSelectionFilter(),
            "Chọn các cây thép (đơn lẻ hoặc nhóm Quantity > 1)");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          return Result.Cancelled;
        }

        if (picks == null || picks.Count == 0)
          return Result.Cancelled;

        // ===== PHÂN LOẠI: Quantity > 1 là nhóm, Quantity = 1 là đơn lẻ =====
        var multiBarIds = new List<ElementId>();
        var singleBarIds = new List<ElementId>();

        foreach (Reference pick in picks)
        {
          if (doc.GetElement(pick) is Rebar rb)
          {
            if (rb.NumberOfBarPositions > 1)
              multiBarIds.Add(rb.Id);
            else
              singleBarIds.Add(rb.Id);
          }
        }

        if (multiBarIds.Count == 0 && singleBarIds.Count == 0)
        {
          TaskDialog.Show("Thông báo", "Không tìm thấy cây thép hợp lệ.");
          return Result.Cancelled;
        }

        // Key: tên RebarBarType, Value: danh sách ElementId
        var rebarsByType = new Dictionary<string, List<ElementId>>();
        var errors = new List<string>();
        int explodedSets = 0;
        int createdBars = 0;
        int groupsCreated = 0;

        using (Transaction t = new Transaction(doc, "CSS Tách Nhóm Cây Thép"))
        {
          t.Start();

          // Cây thép đơn lẻ: không xóa, không tái tạo, chỉ phân loại theo type
          foreach (ElementId id in singleBarIds)
          {
            try
            {
              if (doc.GetElement(id) is Rebar rb)
                AddToDict(rebarsByType, GetBarTypeName(doc, rb), id);
            }
            catch (Exception ex)
            {
              errors.Add($"Rebar {id}: {ex.Message}");
            }
          }

          // Nhóm cây thép (Quantity > 1): thu thập vị trí, convert về single, copy sang các vị trí còn lại
          foreach (ElementId id in multiBarIds)
          {
            try
            {
              Rebar rb = doc.GetElement(id) as Rebar;
              if (rb == null) continue;

              string typeName = GetBarTypeName(doc, rb);
              List<ElementId> newIds = ExplodeRebarSet(doc, rb, errors);

              if (newIds.Count > 0)
              {
                explodedSets++;
                createdBars += newIds.Count;
                foreach (ElementId nid in newIds)
                  AddToDict(rebarsByType, typeName, nid);
              }
            }
            catch (Exception ex)
            {
              errors.Add($"Rebar set {id}: {ex.Message}");
            }
          }

          // Tạo group mới theo type, đặt tên = tên type, pin lại
          //foreach (var kvp in rebarsByType)
          //{
          //  if (kvp.Value.Count == 0) continue;
          //  try
          //  {
          //    Group newGroup = doc.Create.NewGroup(kvp.Value);
          //    for (int suffix = 0; suffix <= 100; suffix++)
          //    {
          //      try
          //      {
          //        newGroup.GroupType.Name = suffix == 0 ? kvp.Key : $"{kvp.Key}_{suffix}";
          //        break;
          //      }
          //      catch { }
          //    }
          //    newGroup.Pinned = true;
          //    groupsCreated++;
          //  }
          //  catch (Exception ex)
          //  {
          //    errors.Add($"Tạo group '{kvp.Key}': {ex.Message}");
          //  }
          //}

          t.Commit();
        }

        // ===== THÔNG BÁO KẾT QUẢ =====
        var sb = new StringBuilder();
        sb.AppendLine(errors.Count == 0 ? "✅ Thành công!" : $"⚠ Hoàn thành với {errors.Count} lỗi.");
        sb.AppendLine($"• Đã tách: {explodedSets} nhóm → {createdBars} cây thép");
        sb.AppendLine($"• Cây thép đơn lẻ giữ nguyên: {singleBarIds.Count}");
        sb.AppendLine($"• Đã tạo và pin: {groupsCreated} nhóm theo type");

        if (errors.Count > 0)
        {
          sb.AppendLine();
          sb.AppendLine("Chi tiết lỗi:");
          foreach (string err in errors.Take(5))
            sb.AppendLine($"- {err}");
          if (errors.Count > 5)
            sb.AppendLine($"... và {errors.Count - 5} lỗi khác");
        }

        TaskDialog.Show("Kết quả tách nhóm cây thép", sb.ToString());

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
        message = ex.Message;
        return Result.Failed;
      }
    }

    /// <summary>
    /// Tách một Rebar có NumberOfBarPositions > 1 thành từng Rebar đơn.
    /// Logic: thu thập tọa độ tất cả vị trí từ geometry trước,
    /// convert bản gốc về single (vị trí 0), sau đó copy sang các vị trí còn lại.
    /// </summary>
    private List<ElementId> ExplodeRebarSet(Document doc, Rebar rebar, List<string> errors)
    {
      var newIds = new List<ElementId>();

      var accessor = rebar.GetShapeDrivenAccessor();
      if (accessor == null)
      {
        errors.Add($"Rebar {rebar.Id}: không phải shape-driven, bỏ qua");
        return newIds;
      }

      int numBars = rebar.NumberOfBarPositions;

      // ===== BƯỚC 1: Thu thập điểm origin của line đầu tiên từ geometry =====
      var positions = new List<XYZ>();
      XYZ basePoint = null;

      Options geoOptions = new Options
      {
        ComputeReferences = false,
        DetailLevel = ViewDetailLevel.Coarse,
        IncludeNonVisibleObjects = true
      };

      for (int i = 0; i < numBars; i++)
      {
        XYZ pt = null;

        // Lấy transform cho bar position này
        Transform barTransform = accessor.GetBarPositionTransform(i);

        // Lấy geometry từ element
        GeometryElement geoElem = (rebar as Element).get_Geometry(geoOptions);

        if (geoElem != null)
        {
          foreach (GeometryObject geoObj in geoElem)
          {
            Curve curve = null;

            if (geoObj is GeometryInstance geoInst)
            {
              GeometryElement instGeom = geoInst.GetInstanceGeometry();
              foreach (GeometryObject instObj in instGeom)
              {
                if (instObj is Curve c && c.Length > 0)
                {
                  curve = c;
                  break;
                }
              }
            }
            else if (geoObj is Curve c && c.Length > 0)
            {
              curve = c;
            }

            if (curve != null && curve.Length > 0)
            {
              // Lấy origin (điểm đầu) của line đầu tiên
              XYZ origin = curve.GetEndPoint(0);
              pt = barTransform.OfPoint(origin);
              break;
            }
          }
        }

        positions.Add(pt);
        if (i == 0) basePoint = pt;
      }

      if (basePoint == null)
      {
        errors.Add($"Rebar {rebar.Id}: không lấy được geometry line vị trí 0");
        return newIds;
      }

      // ===== BƯỚC 2: Convert bản gốc về single (giữ ở vị trí 0) =====
      accessor.SetLayoutAsSingle();
      newIds.Add(rebar.Id);

      // ===== BƯỚC 3: Copy cây single sang từng vị trí còn lại =====
      for (int i = 1; i < numBars; i++)
      {
        XYZ targetPt = positions[i];
        if (targetPt == null) continue;

        XYZ translation = targetPt.Subtract(basePoint);

        try
        {
          ICollection<ElementId> copied = ElementTransformUtils.CopyElement(
            doc, rebar.Id, translation);

          if (copied != null)
            foreach (ElementId cid in copied)
              newIds.Add(cid);
        }
        catch (Exception ex)
        {
          errors.Add($"Rebar {rebar.Id} copy vị trí {i}: {ex.Message}");
        }
      }

      return newIds;
    }

    private RebarHookOrientation GetHookOrientation(Rebar rebar, int end)
    {
      try
      {
        string paramName = end == 0 ? "Hook Orientation at Start" : "Hook Orientation at End";
        Parameter p = rebar.LookupParameter(paramName);
        if (p != null && p.HasValue)
          return p.AsInteger() == 1 ? RebarHookOrientation.Right : RebarHookOrientation.Left;
      }
      catch { }
      return RebarHookOrientation.Left;
    }

    private string GetBarTypeName(Document doc, Rebar rebar)
    {
      try
      {
        RebarBarType barType = doc.GetElement(rebar.GetTypeId()) as RebarBarType;
        return barType?.Name ?? "Unknown";
      }
      catch
      {
        return "Unknown";
      }
    }

    private void AddToDict(Dictionary<string, List<ElementId>> dict, string key, ElementId id)
    {
      if (!dict.ContainsKey(key))
        dict[key] = new List<ElementId>();
      dict[key].Add(id);
    }
  }

  public class RebarSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element elem) => elem is Rebar;
    public bool AllowReference(Reference reference, XYZ position) => false;
  }
}
