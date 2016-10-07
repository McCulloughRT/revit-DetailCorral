using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ListDetailComponents
{
    [Transaction(TransactionMode.Manual)]
    class Command : IExternalCommand
    {
        public View view { get; set; } // The view to be created
        public double yMax { get; set; } // Tracks the max y coordinate excursion in a single row to prevent overlap
        public XYZ insertionPoint { get; set; } // Default starting point for insertion
        public FamilyInstance family { get; set; } // The family to insert

        /// <summary>
        /// Finds all detail components in use in the model
        /// Creates a new drafting view and places one of each
        /// type of component in a grid based on their bounding boxes
        /// so that no two components overlap.
        /// </summary>
        /// <param name="commandData">boilerplate</param>
        /// <param name="message">boilerplate</param>
        /// <param name="elements">boilerplate</param>
        /// <returns></returns>
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Initialize tracked variables
            insertionPoint = new XYZ(0, 0, 0);
            yMax = 0;

            // Count the number of times this plugins has alreadey been run
            // to prevent a view with the same name from being created;
            int countOfEnumViews = GetElements<ViewDrafting>(doc)
                .Where(i => i.ViewName.StartsWith("_Detail Component Enumeration"))
                .Count();

            // Retrieve all detail component families present in the model
            IList<FamilySymbol> components = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DetailComponents)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            // Get the view family type we want to use as a template
            ElementId viewTypeId = GetElements<ViewFamilyType>(doc)
                .Where(i => i.Name == "Typ Int")
                .ToList()
                .First()
                .Id;

            // Create the new drafting view
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("makeView");
                view = ViewDrafting.Create(doc, viewTypeId);
                view.ViewName = "_Detail Component Enumeration_" + countOfEnumViews.ToString();
                tx.Commit();
            }

            // Iterate through family types adding them to the newly created view
            foreach (FamilySymbol sym in components)
            {
                // Check if any instances of the type exist in the model.
                int instanceCount = GetElements<FamilyInstance>(doc)
                    .Where(i => i.Symbol.Id == sym.Id)
                    .Count();
                // If not, or if the name starts with "00" (break lines), skip them
                if (sym.FamilyName.StartsWith("00") || instanceCount == 0)
                {
                    continue;
                }

                // Otherwise add a new instance of the family type to the view
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Transaction Name");
                    try
                    {
                        family = doc.Create.NewFamilyInstance(insertionPoint, sym, view);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException) // This is thrown for line based families
                    {
                        // Lines based families need an overload of NewFamilyInstance that takes a line
                        // instead of a point.
                        XYZ endpoint = new XYZ(insertionPoint.X + 1, insertionPoint.Y, insertionPoint.Z);
                        Line line = Line.CreateBound(insertionPoint, endpoint);
                        family = doc.Create.NewFamilyInstance(line, sym, view);
                    }
                    tx.Commit();
                }

                // Get the XY bounding box for the component that was just placed
                // and move the insertion point for the next component by its max 
                // X value plus 1. If a max Y excursion is reached, move in the negative
                // Y direction to a new line and start over.
                BoundingBoxXYZ bbox = family.get_BoundingBox(view);
                if (bbox.Max.X > 50)
                {
                    insertionPoint = new XYZ(0, yMax - 5, insertionPoint.Z);
                    yMax = 0;
                }
                else
                {
                    insertionPoint = new XYZ(bbox.Max.X + 1, insertionPoint.Y, insertionPoint.Z);
                    if (bbox.Min.Y < yMax)
                    {
                        yMax = bbox.Min.Y;
                    }
                }
            }

            TaskDialog.Show("Enumerate Components", "New drafting view '" + view.ViewName + "' created!\n\n Find under View Template 'None'.");
            return Result.Succeeded;
        }

        public IEnumerable<T> GetElements<T>(Document document) where T : Element
        {
            FilteredElementCollector collector = new FilteredElementCollector(document);
            collector.OfClass(typeof(T));
            return collector.Cast<T>();
        }
    }
}
