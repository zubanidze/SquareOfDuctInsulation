using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using Autodesk.Revit.DB.Mechanical;


namespace DuctInsulationSq
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiApp.ActiveUIDocument.Document;
            try
            {
                FilteredElementCollector ductCollection = new FilteredElementCollector(doc).OfClass(typeof(Duct));
                FilteredElementCollector ductFittingCollection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting);
                FilteredElementCollector ductInsulationCollection = new FilteredElementCollector(doc).OfClass(typeof(DuctInsulation));
                if (ductCollection.Count() == 0)
                {
                    MessageBox.Show("В данном проекте нет воздуховодов");
                    return Result.Cancelled;
                }
                if (ductFittingCollection.Count() == 0)
                {
                    MessageBox.Show("В данном проекте нет соединительных деталей воздуховодов");
                    return Result.Cancelled;
                }
                Element ins = ductInsulationCollection.FirstOrDefault();
                Parameter param = ins.LookupParameter("Площадь_Изоляции");
                if (null == param)
                {
                    MessageBox.Show("В данном проекте нет параметра \"Площадь_Изоляции\"");
                    return Result.Cancelled;
                }
                Transaction tr = new Transaction(doc, "Площадь изоляции");
                tr.Start();
                foreach (DuctInsulation ductInsulation in ductInsulationCollection)
                {
                    if (doc.GetElement(ductInsulation.HostElementId).Category.Id == Category.GetCategory(doc, BuiltInCategory.OST_DuctCurves).Id)
                    {
                        ductInsulation.LookupParameter("Площадь_Изоляции").Set(doc.GetElement(ductInsulation.HostElementId).LookupParameter("Площадь").AsDouble());
                    }
                    if (doc.GetElement(ductInsulation.HostElementId).Category.Id == Category.GetCategory(doc, BuiltInCategory.OST_DuctFitting).Id)
                    {
                        Guid TYPE_PARAMETER_GUID = new Guid("7e8f7b32-a0ba-453f-940e-3fa1f60880ee");
                        ductInsulation.LookupParameter("Площадь_Изоляции").Set(doc.GetElement(ductInsulation.HostElementId).get_Parameter(TYPE_PARAMETER_GUID).AsDouble());
                    }
                }
                tr.Commit();
                MessageBox.Show("Данные параметра площади изоляции успешно обновлены");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                return Result.Failed;
            }
        }
    }
}

