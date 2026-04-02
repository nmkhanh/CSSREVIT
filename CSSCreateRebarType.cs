using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Document = Autodesk.Revit.DB.Document;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSCreateRebarType : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIDocument uidoc = commandData.Application.ActiveUIDocument;
      Document doc = uidoc.Document;

      try
      {
        // ===== ĐỌC FILE =====
        string filePath = @"D:\CSS\CSSREVIT\Rebar_V638.txt";
        if (!File.Exists(filePath))
        {
          TaskDialog.Show("Lỗi", $"Không tìm thấy file:\n{filePath}");
          return Result.Failed;
        }

        // Parse: bỏ header (dòng đầu), lấy cột 1 (tên) và cột 2 (đường kính dạng D22)
        var rows = File.ReadAllLines(filePath)
        .Skip(1) // bỏ header
          .Where(line => !string.IsNullOrWhiteSpace(line))
          .Select(line =>
     {
       string[] parts = line.Split('\t');
       if (parts.Length < 2)
         return null;
       string typeName = parts[0].Trim();
       string diaRaw = parts[1].Trim(); // D22, D16...
       if (!diaRaw.StartsWith("D", StringComparison.OrdinalIgnoreCase))
         return null;
       if (!int.TryParse(diaRaw.Substring(1), out int dia))
         return null;
       return new
       {
         TypeName = typeName,
         Diameter = dia
       };
     })
     .Where(x => x != null)
          .ToList();

        if (rows.Count == 0)
        {
          TaskDialog.Show("Lỗi", "Không có dữ liệu hợp lệ trong file.");
          return Result.Failed;
        }

        // ===== LẤY TẤT CẢ REBARBARTYPE HIỆN CÓ =====
        var existingTypes = new FilteredElementCollector(doc)
          .OfClass(typeof(RebarBarType))
          .Cast<RebarBarType>()
          .ToList();

        // ===== TRANSACTION =====
        int created = 0;
        int skipped = 0;
        var errors = new List<string>();
        var missingTemplates = new HashSet<int>();

        using (Transaction tx = new Transaction(doc, "Create Rebar Types"))
        {
          tx.Start();

          foreach (var row in rows)
          {
            // Tên type mới: typeName_diameter  (ví dụ: P1_22)
            string newName = $"{row.TypeName}_D{row.Diameter}";

            // Đã tồn tại → bỏ qua
            if (existingTypes.Any(t => t.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
            {
              skipped++;
              continue;
            }

            // Tìm template: CSS16, CSS22... (tên chứa "CSS" + đường kính)
            string templateName = $"CSS{row.Diameter}";
            RebarBarType template = existingTypes
                     .FirstOrDefault(t => t.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase));

            if (template == null)
            {
              missingTemplates.Add(row.Diameter);
              errors.Add($"[{newName}] Không tìm thấy template '{templateName}'");
              continue;
            }

            // Duplicate template → đặt tên mới
            try
            {
              RebarBarType newType = template.Duplicate(newName) as RebarBarType;
              if (newType != null)
              {
                created++;
                existingTypes.Add(newType); // cập nhật cache để tránh trùng trong vòng lặp
              }
              else
              {
                errors.Add($"[{newName}] Duplicate trả về null");
              }
            }
            catch (Exception ex)
            {
              errors.Add($"[{newName}] {ex.Message}");
            }
          }

          tx.Commit();
        }

        // ===== THÔNG BÁO 1 LẦN =====
        var sb = new StringBuilder();
        sb.AppendLine($"✓ Đã tạo:   {created} type");
        sb.AppendLine($"⏭ Đã có:    {skipped} type (bỏ qua)");

        if (errors.Count > 0)
        {
          sb.AppendLine($"✗ Lỗi:      {errors.Count}");

          if (missingTemplates.Count > 0)
          {
            sb.AppendLine($"\nThiếu template: {string.Join(", ", missingTemplates.OrderBy(d => d).Select(d => $"CSS{d}"))}");
          }

          sb.AppendLine("\nChi tiết lỗi (tối đa 10):");
          foreach (string err in errors.Take(10))
            sb.AppendLine($"  - {err}");

          if (errors.Count > 10)
            sb.AppendLine($"  ... và {errors.Count - 10} lỗi khác");
        }

        TaskDialog.Show("Kết quả - Create Rebar Types", sb.ToString());
        return Result.Succeeded;
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
        message = ex.Message;
        return Result.Failed;
      }
    }
  }
}
