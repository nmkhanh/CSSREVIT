using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Color = Autodesk.Revit.DB.Color;

namespace CSSREVIT
{
  [Transaction(TransactionMode.Manual)]
  public class CSSLineName : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      try
      {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // 🔥 DATA từ hình bạn gửi
        var data = File.ReadAllLines(@"D:\CSS\CSSREVIT\Rebar_V638.txt")
          .Skip(1)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
        List<string> rebarNames = new List<string>(data.Select(x => $"{x.Split('\t')[0]}_{x.Split('\t')[1]}"));

        using (Transaction tran = new Transaction(doc, "Create LineStyles"))
        {
          tran.Start();

          CreateLineStyles(doc, rebarNames);

          tran.Commit();
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

    private void CreateLineStyles(Document doc, List<string> names)
    {
      try
      {
        Category linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);

        int index = 0;

        foreach (string name in names.Distinct())
        {
          try
          {
            Category existed = linesCategory.SubCategories
                .Cast<Category>()
                .FirstOrDefault(c => c.Name == name);

            if (existed != null)
              continue;

            Category newSubCat = doc.Settings.Categories.NewSubcategory(linesCategory, name);

            Color color = GenerateColor(index++);
            newSubCat.LineColor = color;

            newSubCat.SetLineWeight(3, GraphicsStyleType.Projection);
          }
          catch (Exception exItem)
          {
            Trace.WriteLine($"Error {name}: {exItem}");
          }
        }
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
      }
    }

    private Color GenerateColor(int i)
    {
      try
      {
        byte r = (byte)((i * 70) % 256);
        byte g = (byte)((i * 130) % 256);
        byte b = (byte)((i * 200) % 256);

        return new Color(r, g, b);
      }
      catch (Exception ex)
      {
        Trace.WriteLine(ex.ToString());
        return new Color(0, 0, 0);
      }
    }
  }
}