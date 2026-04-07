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
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace CSSREVIT
{
  /// <summary>
  /// Group các Rebar có cùng type name thành 1 group và pin lại.
  /// Xử lý trường hợp có dấu "-": ví dụ d1-1, d1-2 sẽ chung group d1
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CSSGroupRebarByType : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // ===== SELECT REBARS =====
        IList<Reference> picks;
        try
        {
          picks = uidoc.Selection.PickObjects(
            ObjectType.Element,
            new RebarSelectionFilter(),
            "Chọn các cây thép để group theo type");
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
          return Result.Cancelled;
        }

        if (picks == null || picks.Count == 0)
          return Result.Cancelled;

        // ===== PHÂN LOẠI THEO TYPE NAME =====
        var rebarsByType = new Dictionary<string, List<ElementId>>();

        foreach (Reference pick in picks)
        {
          if (doc.GetElement(pick) is Rebar rb)
          {
            string typeName = GetRebarTypeName(doc, rb);
            string baseName = GetBaseName(typeName);

            if (!rebarsByType.ContainsKey(baseName))
              rebarsByType[baseName] = new List<ElementId>();

            rebarsByType[baseName].Add(rb.Id);
          }
        }

        if (rebarsByType.Count == 0)
        {
          TaskDialog.Show("Thông báo", "Không tìm thấy cây thép hợp lệ.");
          return Result.Cancelled;
        }

        // ===== TẠO GROUP VÀ PIN =====
        int groupsCreated = 0;
        int totalRebars = 0;
        var errors = new List<string>();

        using (Transaction t = new Transaction(doc, "CSS Group Rebar By Type"))
        {
          t.Start();

          foreach (var kvp in rebarsByType)
          {
            string baseName = kvp.Key;
            List<ElementId> rebarIds = kvp.Value;

            if (rebarIds.Count == 0)
              continue;

            totalRebars += rebarIds.Count;

            try
            {
              Group newGroup = doc.Create.NewGroup(rebarIds);

              // Đặt tên group = base name, thêm suffix nếu trùng
              for (int suffix = 0; suffix <= 100; suffix++)
              {
                try
                {
                  newGroup.GroupType.Name = suffix == 0 ? baseName : $"{baseName}_{suffix}";
                  break;
                }
                catch { }
              }

              // Pin group
              newGroup.Pinned = true;
              groupsCreated++;
            }
            catch (Exception ex)
            {
              errors.Add($"Tạo group '{baseName}': {ex.Message}");
            }
          }

          t.Commit();
        }

        // ===== THÔNG BÁO KẾT QUẢ =====
        var sb = new StringBuilder();
        sb.AppendLine(errors.Count == 0 ? "✅ Thành công!" : $"⚠ Hoàn thành với {errors.Count} lỗi.");
        sb.AppendLine($"• Tổng cây thép: {totalRebars}");
        sb.AppendLine($"• Đã tạo và pin: {groupsCreated} nhóm");

        if (errors.Count > 0)
        {
          sb.AppendLine();
          sb.AppendLine("Chi tiết lỗi:");
          foreach (string err in errors.Take(5))
            sb.AppendLine($"- {err}");
          if (errors.Count > 5)
            sb.AppendLine($"... và {errors.Count - 5} lỗi khác");
        }

        TaskDialog.Show("Kết quả group rebar", sb.ToString());

        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        message = ex.Message;
        return Result.Failed;
      }
    }

    /// <summary>
    /// Lấy base name từ type name.
    /// Ví dụ: "d1-1" → "d1", "d1-2" → "d1", "F2" → "F2"
    /// </summary>
    private string GetBaseName(string typeName)
    {
      if (string.IsNullOrEmpty(typeName))
        return "Unknown";

      if (typeName.Contains("-"))
      {
        int index = typeName.IndexOf("-");
        return typeName.Substring(0, index);
      }

      return typeName;
    }

    private string GetRebarTypeName(Document doc, Rebar rebar)
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
  }
}
