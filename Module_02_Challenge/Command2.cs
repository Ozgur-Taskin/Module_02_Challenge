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

#endregion

namespace Module_02_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Tag Active View Algorithm - Challenge:
            // 1- Collect all elements from active view
            // 2- Create multicategory filters for each view 
            // 3- Get tag family symbol by name (LINQ) and collect them in a dictionary
            // 4- Check view type
            // 5- Filter elements
            // 6- CHeck view type and insert tag
            //      - Select each element one by one from filtered elements
            //          - Get insertion point (element can have a point or curve)
            //          - Call the tag family symbol from dictionary
            //          - Create element reference
            //          - Pick the correct tag and Create Tag ==> Need: DOC, TAG FAMILY SYMBOL ID,
            //      VIEW ID, REFERENCE (element to tag), leader, orientation, INSERTION POINT
            //          - Move tag head if needed
            // 7- Create a counter for tags and report back

            // CURTAIN WALL TYPE? WallKind?

            // Tag all views

            // 1- Collect all elements from active view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            // 2- Create multicategory filters for each view

            List<BuiltInCategory> catListAreaPlan = new List<BuiltInCategory>();
            catListAreaPlan.Add(BuiltInCategory.OST_Areas);

            List<BuiltInCategory> catListCeilingPlan = new List<BuiltInCategory>();
            catListCeilingPlan.Add(BuiltInCategory.OST_LightingFixtures);
            catListCeilingPlan.Add(BuiltInCategory.OST_Rooms);

            List<BuiltInCategory> catListFloorPlan = new List<BuiltInCategory>();
            catListFloorPlan.Add(BuiltInCategory.OST_Doors);
            catListFloorPlan.Add(BuiltInCategory.OST_Furniture);
            catListFloorPlan.Add(BuiltInCategory.OST_Rooms);
            catListFloorPlan.Add(BuiltInCategory.OST_Walls);
            catListFloorPlan.Add(BuiltInCategory.OST_Windows);

            List<BuiltInCategory> catListSection = new List<BuiltInCategory>();
            catListSection.Add(BuiltInCategory.OST_Rooms);


            //ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            //collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // 3- Get tag family symbol by name (LINQ) and collect them in a dictionary
            FamilySymbol curLightFixTag = FilterFamilySymbol("M_Lighting Fixture Tag", doc);
            FamilySymbol curRoomTag = FilterFamilySymbol("M_Room Tag", doc);
            FamilySymbol curDoorTag = FilterFamilySymbol("M_Door Tag", doc);
            FamilySymbol curFurnitureTag = FilterFamilySymbol("M_Furniture Tag", doc);
            FamilySymbol curWallTag = FilterFamilySymbol("M_Wall Tag", doc);
            FamilySymbol curCurtainWallTag = FilterFamilySymbol("M_Curtain Wall Tag", doc);
            FamilySymbol curWindowsTag = FilterFamilySymbol("M_Window Tag", doc);

            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Lighting Fixtures", curLightFixTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Doors", curDoorTag);
            tags.Add("Furniture", curFurnitureTag);
            tags.Add("Walls", curWallTag);
            tags.Add("Curtain Walls", curCurtainWallTag);
            tags.Add("Windows", curWindowsTag);

            // Insert tag
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Tag Elements");

                int counter = 0;

                // 4- Check view type
                ViewType curViewType = curView.ViewType;

                if (curViewType == ViewType.AreaPlan)
                {
                    // 5- Filter elements
                    ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catListAreaPlan);
                    collector.WherePasses(catFilter).WhereElementIsNotElementType();

                    // 6 - Insert tag
                    counter = TagAreaPlan(collector, doc, curView, counter);
                }

                if (curViewType == ViewType.CeilingPlan)
                {
                    // 5- Filter elements
                    ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catListCeilingPlan);
                    collector.WherePasses(catFilter).WhereElementIsNotElementType();

                    // 6 - Insert tag
                    counter = TagElements(collector, tags, doc, curView, counter);
                }

                if (curViewType == ViewType.FloorPlan)
                {
                    // 5- Filter elements
                    ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catListFloorPlan);
                    collector.WherePasses(catFilter).WhereElementIsNotElementType();

                    // 6 - Insert tag
                    counter = TagFloorPlan(collector, tags, doc, curView, counter);
                }

                if (curViewType == ViewType.Section)
                {
                    // 5- Filter elements
                    ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catListSection);
                    collector.WherePasses(catFilter).WhereElementIsNotElementType();

                    // 6 - Insert tag
                    counter = TagSection(collector, tags, doc, curView, counter);
                }

                TaskDialog.Show("Tag", counter + " elements tagged.");

                t.Commit();
            }

            return Result.Succeeded;
        }

        private int TagAreaPlan(FilteredElementCollector col,
            Document doc,
            View curView,
            int counter)
        {
            foreach (Element curElem in col)
            {
                ViewPlan curAreaPlan = curView as ViewPlan;
                Area curArea = curElem as Area;

                XYZ point = GetElementPoint(curElem);

                AreaTag curAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(point.X, point.Y));
                curAreaTag.TagHeadPosition = new XYZ(point.X, point.Y, 0);
                curAreaTag.HasLeader = false;

                counter++;
            }

            return counter;
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

        private int TagFloorPlan(FilteredElementCollector col,
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

                // Check wall kind and change tag
                if(curElem.Category.Name == "Walls")
                {
                    Wall curWall = curElem as Wall;
                    WallType curWallType = curWall.WallType;
                    if(curWallType.Kind == WallKind.Curtain)
                    {
                        curTagType = tags["Curtain Walls"]; ;
                    }
                }

                Reference curRef = new Reference(curElem);

                IndependentTag newTag = IndependentTag.Create(
                    doc, curTagType.Id,
                    curView.Id, curRef, false,
                    TagOrientation.Horizontal,
                    point);

                if (curTagType == tags["Windows"])
                {
                    newTag.TagHeadPosition = new XYZ(point.X, (point.Y + 3), point.Z);
                }

                counter++;
            }

            return counter;
        }

        private int TagSection(FilteredElementCollector col,
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

                if(curTagType == tags["Rooms"])
                {
                    newTag.TagHeadPosition = new XYZ(point.X, point.Y, (point.Z + 3));
                }

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
                .Where(x => x.FamilyName.Equals(familyName))
                .First();
            return familySymbol;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
