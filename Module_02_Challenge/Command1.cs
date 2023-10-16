#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Documents;

#endregion

namespace Module_02_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // One Tag in Active View Algorithm:
            // 1- Collect all elements from active view
            // 2- Create multicategory filter and filter elements
            // 3- Get tag family symbol by name (LINQ) and collect them in a dictionary
            // 4- Insert tag
            //      - Select each element one by one from filtered elements
            //          - Get insertion point (element can have a point or curve)
            //          - Call the tag family symbol from dictionary
            //          - Create element reference
            // 5- Pick the correct tag and Create Tag ==> Need: DOC, TAG FAMILY SYMBOL ID,
            //      VIEW ID, REFERENCE (element to tag), leader, orientation, INSERTION POINT
            // 6- Create a counter for tags and report back


            // 1- Collect all elements from active view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            // 2- Create multicategory filter and filter elements
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Doors);

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // 3- Get tag family symbol by name (LINQ) and collect them in a dictionary
            FamilySymbol curDoorTag = FilterFamilySymbol("M_Door Tag", doc);

            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Doors", curDoorTag);

            // 4 - Insert tag
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Tag Elements");

                int counter = 0;

                counter = TagElements(collector, tags, doc, curView, counter);
                               

                TaskDialog.Show("Tag", counter + " elements tagged.");

                t.Commit();
            }

            return Result.Succeeded;
        }

        private int TagElements(FilteredElementCollector col, 
            Dictionary<string, FamilySymbol> tags,
            Document doc,
            View curView,
            int counter)
        {
            foreach (Element curElem in col)
            {
                Location curLoc = curElem.Location;

                if (curLoc == null)
                {
                    continue;
                }

                XYZ point = GetElementPoint(curElem);
                FamilySymbol curTagType = tags[curElem.Category.Name];
                Reference curRef = new Reference(curElem);

                IndependentTag newTag = IndependentTag.Create(
                    doc, curTagType.Id,
                    curView.Id, curRef, false,
                    TagOrientation.Horizontal,
                    point);

                counter++;
            }

            return counter;
        }

        private XYZ GetElementPoint(Element curElem)
        {
            XYZ insPoint;
            LocationPoint locPoint;
            LocationCurve locCurve;

            Location curLoc = curElem.Location;

            locPoint = curLoc as LocationPoint;
            if (locPoint != null)
            {
                insPoint = locPoint.Point;
            }
            else
            {
                locCurve = curLoc as LocationCurve;
                Curve curCurve = locCurve.Curve;
                insPoint = GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
            }

            return insPoint;
        }

        private XYZ GetMidpointBetweenTwoPoints(XYZ p1, XYZ p2)
        {
            XYZ mid = new XYZ(
                (p1.X + p2.X) / 2,
                (p1.Y + p2.Y) / 2,
                (p1.Z + p2.Z) / 2);

            return mid;
        }

        private FamilySymbol FilterFamilySymbol(string familyName, Document doc)
        {
            FamilySymbol familySymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals (familyName))
                .First();
            return familySymbol;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
