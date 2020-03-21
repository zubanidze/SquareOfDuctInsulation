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
        public static double CalculateAreaOfFitting(FamilyInstance element, Document doc)
        {
            Options options = new Options
            {
                View = doc.ActiveView
            };
            double area = 0;
            var geometryElement = element.get_Geometry(options);
            var mechEl = element.MEPModel as MechanicalFitting;
            foreach ( var geomObj in geometryElement)
            {
                if (geomObj is GeometryInstance)
                {
                    var geomInstance = geomObj as GeometryInstance;
                    var geomEl = geomInstance.GetInstanceGeometry();

                    foreach (var ge in geomEl)
                    {
                        if (!(ge is Solid)) continue;

                        area += (ge as Solid).SurfaceArea;
                    }
                }
            }
            
            foreach (Connector connector in mechEl.ConnectorManager.Connectors)
            {
                if (connector.Shape == ConnectorProfileType.Rectangular)
                {
                    area -= connector.Width * connector.Height;
                }

                else if (connector.Shape == ConnectorProfileType.Round)
                    area -= Math.PI * Math.Pow(connector.Radius, 2);
            }

            return area;
        }
        public Result SetAreasOfInsulations (Document doc)
        {
            try
            {
                FilteredElementCollector ductCollection = new FilteredElementCollector(doc).OfClass(typeof(Duct));
                FilteredElementCollector ductFittingCollection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting);
                FilteredElementCollector ductInsulationCollection = new FilteredElementCollector(doc).OfClass(typeof(DuctInsulation));
                if (ductCollection.Count() == 0)
                {
                    TaskDialog.Show("Ошибка", "В данном проекте нет воздуховодов");
                    return Result.Cancelled;
                }
                if (ductFittingCollection.Count() == 0)
                {
                    TaskDialog.Show("Ошибка", "В данном проекте нет соединительных деталей воздуховодов");
                    return Result.Cancelled;
                }
                Element ins = ductInsulationCollection.FirstOrDefault();
                Parameter param = ins.LookupParameter("Площадь_Изоляции");
                if (null == param)
                {
                    TaskDialog.Show("Ошибка", "В данном проекте нет параметра \"Площадь_Изоляции\"");
                    return Result.Cancelled;
                }
                Transaction tr = new Transaction(doc, "Площадь изоляции");
                tr.Start();
                foreach (DuctInsulation ductInsulation in ductInsulationCollection)
                {
                    var hostElement = doc.GetElement(ductInsulation.HostElementId);
                    var hostElementBIC = (BuiltInCategory)doc.GetElement(ductInsulation.HostElementId).Category.Id.IntegerValue;
                    if (hostElementBIC == BuiltInCategory.OST_DuctCurves)
                    {
                        ductInsulation.LookupParameter("Площадь_Изоляции").Set(hostElement.LookupParameter("Площадь").AsDouble());
                    }
                    if (hostElementBIC == BuiltInCategory.OST_DuctFitting)
                    {
                        double area = CalculateAreaOfFitting((hostElement as FamilyInstance), doc);
                        ductInsulation.LookupParameter("Площадь_Изоляции").Set(area);
                    }
                }
                tr.Commit();
                TaskDialog.Show("Готово", "Данные параметра площади изоляции успешно обновлены");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
                return Result.Failed;
            }
        }
        public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)

        { 
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            return SetAreasOfInsulations(doc);
            
        }
    }
}

