using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Text.RegularExpressions;
//using Autodesk.Revit.Creation;

namespace DbElement
{
 

    [Transaction(TransactionMode.Manual)]

    public class DBElement : IExternalCommand
    {

        public class MidasPoint
        {
            public int ID { get; set; }
            public XYZ Loc { get; set; }

            public MidasPoint(int PointID, XYZ Point)
            {
                ID = PointID;
                Loc = new XYZ(UnitUtils.ConvertFromInternalUnits(Point.X, DisplayUnitType.DUT_METERS),
                                UnitUtils.ConvertFromInternalUnits(Point.Y, DisplayUnitType.DUT_METERS),
                                UnitUtils.ConvertFromInternalUnits(Point.Z, DisplayUnitType.DUT_METERS));
            }
        }

        public class MidasElement
        {

            public int ID { get; set; }
            public int RevitID { get; set; }
            public int StartPoint { get; set; }
            public int EndPoint { get; set; }
            public int SectionNumber { get; set; }
            public int MaterialNumber { get; set; }

            public MidasElement(int iD, int revitID, int startPoint, int endPoint, int sectionNumber, int materialNumber)
            {
                ID = iD;
                RevitID = revitID;
                StartPoint = startPoint;
                EndPoint = endPoint;
                SectionNumber = sectionNumber;
                MaterialNumber = materialNumber;
            }
        }

        public class PointBoundary
        {
            public int ID { get; set; }
            public int PointNumber { get; set; }
            public int DX { get; set; }
            public int DY { get; set; }
            public int DZ { get; set; }
            public int RX { get; set; }
            public int RY { get; set; }
            public int RZ { get; set; }

            public PointBoundary(int iD, int pointNumber, Element elem)
            {
                ID = iD;
                PointNumber = pointNumber;
                DX = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_X).AsInteger()-1);
                DY = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_Y).AsInteger()-1);
                DZ = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_Z).AsInteger()-1);

                RX = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_ROT_X).AsInteger() - 1);
                RY = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_ROT_Y).AsInteger() - 1);
                RZ = Math.Abs(elem.get_Parameter(BuiltInParameter.BOUNDARY_DIRECTION_ROT_Z).AsInteger() - 1);
            }
        }

        public class MidasPointLoad
        {
            public int ID { get; set; }
            public int PointNumber { get; set; }
            public string LoadCase { get; set; }
            public double FX { get; set; }
            public double FY { get; set; }
            public double FZ { get; set; }
            public double MX { get; set; }
            public double MY { get; set; }
            public double MZ { get; set; }

            public MidasPointLoad(int iD, int pointNumber, XYZ forceVector, XYZ momentVector, string loadCase)
            {
                ID = iD;
                PointNumber = pointNumber;

                LoadCase = loadCase;

                FX = UnitUtils.ConvertFromInternalUnits(forceVector.X, DisplayUnitType.DUT_KILONEWTONS);
                FY = UnitUtils.ConvertFromInternalUnits(forceVector.Y, DisplayUnitType.DUT_KILONEWTONS);
                FZ = UnitUtils.ConvertFromInternalUnits(forceVector.Z, DisplayUnitType.DUT_KILONEWTONS);

                MX = UnitUtils.ConvertFromInternalUnits(momentVector.X, DisplayUnitType.DUT_KILONEWTON_METERS);
                MY = UnitUtils.ConvertFromInternalUnits(momentVector.Y, DisplayUnitType.DUT_KILONEWTON_METERS);
                MZ = UnitUtils.ConvertFromInternalUnits(momentVector.Z, DisplayUnitType.DUT_KILONEWTON_METERS);
            }
        }

        public class MidasLoadCase
        {
            public string LoadCaseName { get; set; }
            public string LoadCategoryName { get; set; }
            public string LoadNature { get; set; }

            public MidasLoadCase(string loadCaseName, string loadCategoryName, string loadNature)
            {
                LoadCaseName = loadCaseName;
                LoadCategoryName = loadCategoryName;
                LoadNature = loadNature;
            }
        }

        public class MidasElementLoad
        {
            public int ID { get; set; }
            public int HostElementID { get; set; }
            public string LoadCase { get; set; }
            public double StartPoint { get; set; }
            public double EndPoint { get; set; }
            public double FX1 { get; set; }
            public double FX2 { get; set; }
            public double FY1 { get; set; }
            public double FY2 { get; set; }
            public double FZ1 { get; set; }
            public double FZ2 { get; set; }
            public double MX1 { get; set; }
            public double MX2 { get; set; }
            public double MY1 { get; set; }
            public double MY2 { get; set; }
            public double MZ1 { get; set; }
            public double MZ2 { get; set; }

            public MidasElementLoad(int iD, int hostElementID, string loadCase, double startPoint, double endPoint, LineLoad loadForces)
            {
                ID = iD;
                HostElementID = hostElementID;
                LoadCase = loadCase;
                StartPoint = startPoint;
                EndPoint = endPoint;
                XYZ forceVector1 = loadForces.ForceVector1;
                XYZ forceVector2 = loadForces.ForceVector2;
                XYZ momentVector1 = loadForces.MomentVector1;
                XYZ momentVector2 = loadForces.MomentVector2;
                FX1 = UnitUtils.ConvertFromInternalUnits(forceVector1.X, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FX2 = UnitUtils.ConvertFromInternalUnits(forceVector2.X, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FY1 = UnitUtils.ConvertFromInternalUnits(forceVector1.Y, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FY2 = UnitUtils.ConvertFromInternalUnits(forceVector2.Y, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FZ1 = UnitUtils.ConvertFromInternalUnits(forceVector1.Z, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FZ2 = UnitUtils.ConvertFromInternalUnits(forceVector2.Z, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                MX1 = UnitUtils.ConvertFromInternalUnits(momentVector1.X, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MX2 = UnitUtils.ConvertFromInternalUnits(momentVector2.X, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MY1 = UnitUtils.ConvertFromInternalUnits(momentVector1.Y, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MY2 = UnitUtils.ConvertFromInternalUnits(momentVector2.Y, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MZ1 = UnitUtils.ConvertFromInternalUnits(momentVector1.Z, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MZ2 = UnitUtils.ConvertFromInternalUnits(momentVector2.Z, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
            }
            public MidasElementLoad(int iD, int hostElementID, double startPoint, double endPoint, MidasLineLoad lineLoad)
            {
                ID = iD;
                HostElementID = hostElementID;
                LoadCase = lineLoad.LoadCase;
                StartPoint = startPoint;
                EndPoint = endPoint;
                FX1 = lineLoad.FX1;
                FX2 = lineLoad.FX2;
                FY1 = lineLoad.FY1;
                FY2 = lineLoad.FY2;
                FZ1 = lineLoad.FZ1;
                FZ2 = lineLoad.FZ2;
                MX1 = lineLoad.MX1;
                MX2 = lineLoad.MX2;
                MY1 = lineLoad.MY1;
                MY2 = lineLoad.MY2;
                MZ1 = lineLoad.MZ1;
                MZ2 = lineLoad.MZ2;
            }
        }

        public class MidasLineLoad
        {
            public int ID { get; set; }
            public string LoadCase { get; set; }
            public XYZ StartPoint { get; set; }
            public XYZ EndPoint { get; set; }
            public double FX1 { get; set; }
            public double FX2 { get; set; }
            public double FY1 { get; set; }
            public double FY2 { get; set; }
            public double FZ1 { get; set; }
            public double FZ2 { get; set; }
            public double MX1 { get; set; }
            public double MX2 { get; set; }
            public double MY1 { get; set; }
            public double MY2 { get; set; }
            public double MZ1 { get; set; }
            public double MZ2 { get; set; }

            public MidasLineLoad(string loadCase, XYZ startPoint, XYZ endPoint, LineLoad loadForces)
            {
                LoadCase = loadCase;
                StartPoint = new XYZ(UnitUtils.ConvertFromInternalUnits(startPoint.X, DisplayUnitType.DUT_METERS), 
                                    UnitUtils.ConvertFromInternalUnits(startPoint.Y, DisplayUnitType.DUT_METERS), 
                                    UnitUtils.ConvertFromInternalUnits(startPoint.Z, DisplayUnitType.DUT_METERS));
                EndPoint = new XYZ(UnitUtils.ConvertFromInternalUnits(endPoint.X, DisplayUnitType.DUT_METERS),
                                    UnitUtils.ConvertFromInternalUnits(endPoint.Y, DisplayUnitType.DUT_METERS),
                                    UnitUtils.ConvertFromInternalUnits(endPoint.Z, DisplayUnitType.DUT_METERS));
                XYZ forceVector1 = loadForces.ForceVector1;
                XYZ forceVector2 = loadForces.ForceVector2;
                XYZ momentVector1 = loadForces.MomentVector1;
                XYZ momentVector2 = loadForces.MomentVector2;
                FX1 = UnitUtils.ConvertFromInternalUnits(forceVector1.X, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FX2 = UnitUtils.ConvertFromInternalUnits(forceVector2.X, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FY1 = UnitUtils.ConvertFromInternalUnits(forceVector1.Y, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FY2 = UnitUtils.ConvertFromInternalUnits(forceVector2.Y, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FZ1 = UnitUtils.ConvertFromInternalUnits(forceVector1.Z, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                FZ2 = UnitUtils.ConvertFromInternalUnits(forceVector2.Z, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                MX1 = UnitUtils.ConvertFromInternalUnits(momentVector1.X, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MX2 = UnitUtils.ConvertFromInternalUnits(momentVector2.X, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MY1 = UnitUtils.ConvertFromInternalUnits(momentVector1.Y, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MY2 = UnitUtils.ConvertFromInternalUnits(momentVector2.Y, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MZ1 = UnitUtils.ConvertFromInternalUnits(momentVector1.Z, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                MZ2 = UnitUtils.ConvertFromInternalUnits(momentVector2.Z, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
            }
        }

        //Variables utiles (app et doc)

        Application m_rvtApp;
        Document m_rvtDoc;

        public FilteredElementCollector GetStructuralElements(Document doc)
        {
            // what categories of family instances
            // are we interested in?

            BuiltInCategory[] bics = new BuiltInCategory[] {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation
            };

            IList<ElementFilter> a
              = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics)
            {
                a.Add(new ElementCategoryFilter(bic));
            }

            LogicalOrFilter categoryFilter
              = new LogicalOrFilter(a);

            LogicalAndFilter familyInstanceFilter
              = new LogicalAndFilter(categoryFilter,
                new ElementClassFilter(
                  typeof(FamilyInstance)));

            IList<ElementFilter> b
              = new List<ElementFilter>(6);

            b.Add(new ElementClassFilter(
              typeof(Wall)));

            b.Add(new ElementClassFilter(
              typeof(Floor)));

            //b.Add(new ElementClassFilter(
            //  typeof(ContFooting)));

            b.Add(familyInstanceFilter);

            LogicalOrFilter classFilter
              = new LogicalOrFilter(b);

            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.WherePasses(classFilter);

            return collector;
        }

        public FilteredElementCollector GetSupportAndLoads(Document doc)
        {
            // what categories of family instances
            // are we interested in?

            BuiltInCategory[] bics = new BuiltInCategory[] {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation
            };

            IList<ElementFilter> b
              = new List<ElementFilter>(6);

            b.Add(new ElementClassFilter(
              typeof(PointLoad)));

            b.Add(new ElementClassFilter(
              typeof(LineLoad)));

            b.Add(new ElementClassFilter(
              typeof(AreaLoad)));

            b.Add(new ElementClassFilter(
                typeof(BoundaryConditions)));

            LogicalOrFilter classFilter
              = new LogicalOrFilter(b);

            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.WherePasses(classFilter);

            return collector;
        }

        public bool PointOnLine(XYZ P, XYZ A, XYZ B)
        {
            double D = Math.Abs(A.DistanceTo(B));
            double D1 = Math.Abs(P.DistanceTo(A));
            double D2 = Math.Abs(P.DistanceTo(B));
            if (Math.Abs(D - (D1 + D2)) < 0.0000000001)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsColin(XYZ A, XYZ B, XYZ P, XYZ Q)
        {
            double X1 = A.X - B.X, Y1 = A.Y - B.Y, Z1 = A.Z - B.Z, X2 = P.X - Q.X, Y2 = P.Y - Q.Y, Z2 = P.Z - Q.Z;
            if (Math.Abs(X1*Y2-X2*Y1)<0.0000000001 && Math.Abs(X1*Z2 - Z1*X2) < 0.0000000001 && Math.Abs(Y1 * Z2 - Z1 * Y2) < 0.0000000001)
            {
                return true;
            }
            return false;
        }

        public void getMidasParams(Element elem, ref int nElements, ref int nNodes, ref List<MidasPoint> Nodes, ref List<MidasElement> Elements, ref List<string> MaterialList, ref List<string> SectionList, ref List<int> ElementNodes)
        {
            Location loc = elem.Location;

            Regex steel = new Regex(@"(S[0-9]{3,4})");

            if (loc is LocationCurve)
            {
                LocationCurve locCurve = (LocationCurve)loc;
                Curve crv = locCurve.Curve;

                XYZ pt = crv.GetEndPoint(0);

                int startPoint, endPoint;

                var where = Nodes.Find(i => Math.Abs(i.Loc.X - UnitUtils.ConvertFromInternalUnits(pt.X, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Y - UnitUtils.ConvertFromInternalUnits(pt.Y, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Z - UnitUtils.ConvertFromInternalUnits(pt.Z, DisplayUnitType.DUT_METERS)) < 0.0000000001);

                if (where == null)
                {
                    MidasPoint P1 = new MidasPoint(nNodes, pt);
                    Nodes.Add(P1);
                    startPoint = nNodes;
                    ElementNodes.Add(nNodes);
                    nNodes++;
                }
                else
                {
                    startPoint = where.ID;
                }

                pt = crv.GetEndPoint(1);

                where = Nodes.Find(i => Math.Abs(i.Loc.X - UnitUtils.ConvertFromInternalUnits(pt.X, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Y - UnitUtils.ConvertFromInternalUnits(pt.Y, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Z - UnitUtils.ConvertFromInternalUnits(pt.Z, DisplayUnitType.DUT_METERS)) < 0.0000000001);

                if (where == null)
                {
                    MidasPoint P2 = new MidasPoint(nNodes, pt);
                    Nodes.Add(P2);
                    endPoint = nNodes;
                    ElementNodes.Add(nNodes);
                    nNodes++;
                }
                else
                {
                    endPoint = where.ID;
                }

                Parameter section;
                Parameter materiau = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
                int MaterialNumber;


                Match match = steel.Match(materiau.AsValueString());

                if (match.Success)
                {
                    string nuance_acier = match.Groups[1].Value;

                    if (MaterialList.IndexOf(nuance_acier) == -1)
                    {
                        MaterialList.Add(nuance_acier);
                    }
                    MaterialNumber = MaterialList.IndexOf(nuance_acier) + 1;
                    section = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                }
                else if (materiau.AsValueString().IndexOf("béton") != -1)
                {
                    string beton = "C30/37";
                    if (MaterialList.IndexOf(beton) == -1)
                    {
                        MaterialList.Add(beton);
                    }
                    MaterialNumber = MaterialList.IndexOf(beton) + 1;
                    section = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                }
                else
                {
                    string beton = "C30/37";
                    if (MaterialList.IndexOf(beton) == -1)
                    {
                        MaterialList.Add(beton);
                    }
                    MaterialNumber = MaterialList.IndexOf(beton) + 1;
                    section = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                }


                if (SectionList.IndexOf(section.AsValueString()) ==-1)
                {
                    SectionList.Add(section.AsValueString());
                }

                int SectionNumber = SectionList.IndexOf(section.AsValueString())+1;

                MidasElement element = new MidasElement(nElements, elem.Id.IntegerValue, startPoint, endPoint, SectionNumber, MaterialNumber);

                nElements += 1;

                Elements.Add(element);

            }

            if (loc is LocationPoint)
            {

            }
        }

        public void getBoundaryConditions(Element elem, ref int nNodes, ref int nBoundaries, ref List<MidasPoint> Nodes,ref List<PointBoundary> Boundaries)
        {
            BoundaryConditions con = (BoundaryConditions)elem;
            XYZ pt = con.Point;

            var where = Nodes.Find(i => Math.Abs(i.Loc.X - UnitUtils.ConvertFromInternalUnits(pt.X, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Y - UnitUtils.ConvertFromInternalUnits(pt.Y, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                    && Math.Abs(i.Loc.Z - UnitUtils.ConvertFromInternalUnits(pt.Z, DisplayUnitType.DUT_METERS)) < 0.0000000001);

            if (where == null)
            {
                MidasPoint point = new MidasPoint(nNodes, pt);
                PointBoundary support = new PointBoundary(nBoundaries, nNodes, elem);

                nNodes++;
                nBoundaries++;

                Boundaries.Add(support);
                Nodes.Add(point);
            }

            else
            {
                    PointBoundary support = new PointBoundary(nBoundaries, where.ID, elem);
                    nBoundaries++;
                    Boundaries.Add(support);
            }
        }

        public void getPointLoad(Element elem, ref int nNodes, ref int nPointLoads, ref List<MidasPoint> Nodes, ref List<MidasPointLoad> PointLoads, ref List<MidasLoadCase> LoadCases)
        {
            PointLoad load = (PointLoad)elem;
            XYZ pt = load.Point;

            LoadBase loadBase = (LoadBase)elem;
            string loadCaseName = loadBase.LoadCaseName;

            var where = Nodes.Find(i => Math.Abs(i.Loc.X - UnitUtils.ConvertFromInternalUnits(pt.X, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                     && Math.Abs(i.Loc.Y - UnitUtils.ConvertFromInternalUnits(pt.Y, DisplayUnitType.DUT_METERS)) < 0.0000000001
                                     && Math.Abs(i.Loc.Z - UnitUtils.ConvertFromInternalUnits(pt.Z, DisplayUnitType.DUT_METERS)) < 0.0000000001);

            if (where == null)
            {
                MidasPoint point = new MidasPoint(nNodes, pt);
                MidasPointLoad pointLoad = new MidasPointLoad(nPointLoads, nNodes, load.ForceVector, load.MomentVector, loadCaseName);

                nNodes++;
                nPointLoads++;

                PointLoads.Add(pointLoad);
                Nodes.Add(point);
            }

            else
            {
                MidasPointLoad pointLoad = new MidasPointLoad(nPointLoads, where.ID, load.ForceVector, load.MomentVector, loadCaseName);
                nPointLoads++;
                PointLoads.Add(pointLoad);
            }


            if (LoadCases.Find(i => i.LoadCaseName == loadCaseName) == null)
            {
                MidasLoadCase loadCase = new MidasLoadCase(loadCaseName, loadBase.LoadCategoryName, loadBase.LoadNatureName);
                LoadCases.Add(loadCase);
            }
        }

        public void getLineLoad(Element elem, ref int nElementLoads, ref List<MidasElementLoad> ElementLoads, ref List<MidasLineLoad> LineLoads, ref List<MidasLoadCase> LoadCases)
        {
            LoadBase loadBase = (LoadBase)elem;
            LineLoad loadForces = (LineLoad)elem;
            string loadCaseName = loadBase.LoadCaseName;
            Element hostElement = loadBase.HostElement;
            AnalyticalModel model = (AnalyticalModel)hostElement;



            if (LoadCases.Find(i => i.LoadCaseName == loadCaseName) == null)
            {
                MidasLoadCase loadCase = new MidasLoadCase(loadCaseName, loadBase.LoadCategoryName, loadBase.LoadNatureName);
                LoadCases.Add(loadCase);
            }

            if (loadBase.IsHosted == true)
            {
                MidasElementLoad load = new MidasElementLoad(nElementLoads, model.GetElementId().IntegerValue, loadCaseName, 0, 1, loadForces);
                ElementLoads.Add(load);
                nElementLoads++;
            }

            else
            {
                MidasLineLoad load = new MidasLineLoad(loadCaseName, loadForces.StartPoint, loadForces.EndPoint, loadForces);
                LineLoads.Add(load);
            }
        }

        public string MidasGeometryString(List<MidasPoint> points)
        {
            string s="";
            foreach(var point in points)
            {
                s += point.ID.ToString() + ", "+ point.Loc.X.ToString().Replace(",", ".") + ", " + point.Loc.Y.ToString().Replace(",", ".") + ", " + point.Loc.Z.ToString().Replace(",", ".") + "\n";
            }
            return s;
        }

        public string MidasElementString(List<MidasElement> elements)
        {
            string s = "";
            foreach(var element in elements)
            {
                s += element.ID.ToString() + ", BEAM," + element.MaterialNumber.ToString() + ", " + element.SectionNumber.ToString() + ", " + 
                    element.StartPoint.ToString() + ", " + element.EndPoint.ToString() + ", 0, 0 \n";
            }
            return s;
        }

        public string MidasBoundaryString(List<PointBoundary> supports)
        {
            string s = "";
            foreach(var support in supports)
            {
                s += support.PointNumber.ToString() + ", " + support.DX.ToString() + support.DY.ToString() + support.DZ.ToString() + support.RX.ToString() + support.RY.ToString() + support.RZ.ToString() + ",\n";
            }
            return s;
        }

        public List<string> MidasLoadCaseString(List<MidasLoadCase> loadCases)
        {
            List<string> s = new List<string>();
            List<string> t = new List<string>();
            string s1 = "";
            foreach (var loadCase in loadCases)
            {
                switch (loadCase.LoadNature)
                {
                    case "Permanente":
                        s1 = "D";
                        break;
                    case "Exploitation":
                        s1 = "L";
                        break;
                }
                string loadCaseName = loadCase.LoadCaseName;

                s.Add(loadCaseName + ", " + s1);
                t.Add("*USE-STLD, " + loadCaseName);
                t.Add("; EndOf Data for Load Case " + loadCaseName);
            }
            return s.Concat(t).ToList();
        }

        public void writePointLoadString(MidasPointLoad pointLoad, ref List<string> Lines)
        {
            string s = "*CONLOAD \n" + pointLoad.PointNumber.ToString().Replace(",",".") + ", " + pointLoad.FX.ToString().Replace(",", ".") + ", " + pointLoad.FY.ToString().Replace(",", ".") + ", " + pointLoad.FZ.ToString().Replace(",", ".") + ", "
                 + pointLoad.MX.ToString().Replace(",", ".") + ", " + pointLoad.MY.ToString().Replace(",", ".") + ", " + pointLoad.MZ.ToString().Replace(",", ".") + ", \n";
            Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + pointLoad.LoadCase), s);
        }

        public void writeElementLoadString(MidasElementLoad elementLoad, ref List<string> Lines, List<MidasElement> Elements)
        {
            int ID = elementLoad.HostElementID;
            string loadCaseName = elementLoad.LoadCase;
            double fX1 = elementLoad.FX1, fX2 = elementLoad.FX2, fY1 = elementLoad.FY1, fY2 = elementLoad.FY2, fZ1 = elementLoad.FZ1, fZ2 = elementLoad.FZ2;
            double mX1 = elementLoad.MX1, mX2 = elementLoad.MX2, mY1 = elementLoad.MY1, mY2 = elementLoad.MY2, mZ1 = elementLoad.MZ1, mZ2 = elementLoad.MZ2;

            string startpoint = elementLoad.StartPoint.ToString().Replace(",", ".");
            string endpoint = elementLoad.EndPoint.ToString().Replace(",", ".");


            Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), "*BEAMLOAD \n");

            foreach (var element in Elements)
            {
                if (element.RevitID == ID)
                {
                    string id = element.ID.ToString();
                    if (fX1 != 0 || fX2 != 0)
                    {
                        string s = id + ", LINE, UNILOAD, GX, NO , NO, aDir[1], , , , " + startpoint + ",  " + fX1.ToString().Replace(",",".") + ", " + endpoint + ", " + fX2.ToString().Replace(",", ".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                    if (fY1 != 0 || fY2 != 0)
                    {
                        string s = id + ", LINE, UNILOAD, GY, NO , NO, aDir[1], , , , " + startpoint + ",  " + fY1.ToString().Replace(",", ".") + ", " + endpoint + ", " + fY2.ToString().Replace(",", ".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                    if (fZ1 != 0 || fZ2 != 0)
                    {
                        string s = id + ", LINE, UNILOAD, GZ, NO , NO, aDir[1], , , , " + startpoint + ",  " + fZ1.ToString().Replace(",", ".") + ", " + endpoint + ", " + fZ2.ToString().Replace(",", ".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                    if (mX1 != 0 || mX2 != 0)
                    {
                        string s = id + ", LINE, UNIMOMENT, GX, NO , NO, aDir[1], , , , " + startpoint + ",  " + mX1.ToString().Replace(",", ".") + ", " + endpoint + ", " + mX2.ToString().Replace(",",".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                    if (mY1 != 0 || mY2 != 0)
                    {
                        string s = id + ", LINE, UNIMOMENT, GY, NO , NO, aDir[1], , , , " + startpoint + ",  " + mY1.ToString().Replace(",", ".") + ", " + endpoint + ", " + mY2.ToString().Replace(",", ".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                    if (mZ1 != 0 || mZ2 != 0)
                    {
                        string s = id + ", LINE, UNIMOMENT, GZ, NO , NO, aDir[1], , , , " + startpoint + ",  " + mZ1.ToString().Replace(",", ".") + ", " + endpoint + ", " + mZ2.ToString().Replace(",",".") + ", 0, 0, 0, 0, , NO, 0, 0, NO, \n";
                        Lines.Insert(Lines.IndexOf("; EndOf Data for Load Case " + loadCaseName), s);
                    }
                }
            }
        }

        public void LineToElementLoad(MidasLineLoad lineLoad, ref List<MidasElementLoad> ElementLoads, ref int nElementLoads, List<MidasPoint> Nodes, List<MidasElement> Elements)
        {
            XYZ loadStartPoint = lineLoad.StartPoint;
            XYZ loadEndPoint = lineLoad.EndPoint;
            string loadCase = lineLoad.LoadCase;

            foreach (var element in Elements)
            {
                XYZ startPoint = Nodes[element.StartPoint - 1].Loc;
                XYZ endPoint = Nodes[element.EndPoint - 1].Loc;

                if (IsColin(startPoint, endPoint, loadStartPoint, loadEndPoint))
                {
                    bool b1 = PointOnLine(startPoint, loadStartPoint, loadEndPoint), b2 = PointOnLine(endPoint, loadStartPoint, loadEndPoint);

                    if (b1 && b2)
                    {
                        MidasElementLoad elementLoad = new MidasElementLoad(nElementLoads, element.RevitID, 0, 1, lineLoad);
                        ElementLoads.Add(elementLoad);
                        nElementLoads++;
                    }
                    else if (b1 || b2)
                    {
                        double start, end;
                        if (b1)
                        {
                            start = 0;
                            if (PointOnLine(loadStartPoint, startPoint, endPoint))
                            {
                                end = startPoint.DistanceTo(loadStartPoint) / startPoint.DistanceTo(endPoint);
                            }
                            else
                            {
                                end = startPoint.DistanceTo(loadEndPoint) / startPoint.DistanceTo(endPoint);
                            }
                        }
                        else
                        {
                            end = 1;
                            if (PointOnLine(loadStartPoint, startPoint, endPoint))
                            {
                                start = startPoint.DistanceTo(loadStartPoint) / startPoint.DistanceTo(endPoint);
                            }
                            else
                            {
                                start = startPoint.DistanceTo(loadEndPoint) / startPoint.DistanceTo(endPoint);
                            }
                        }

                        MidasElementLoad elementLoad = new MidasElementLoad(nElementLoads, element.RevitID, start, end, lineLoad);
                        ElementLoads.Add(elementLoad);
                        nElementLoads++;
                    }
                } 
            }
        }

        public void splitElement(int N, ref List<MidasElement> Elements, ref int nElements, List<MidasPoint> Nodes, int length)
        {
            bool stop = false;
            int i = 0;
            XYZ P1;
            XYZ P2;
            MidasElement currentElement = Elements[0];

            XYZ P = Nodes[N - 1].Loc;
            while (!stop && i < length)
            {
                currentElement = Elements[i];
                int n1 = currentElement.StartPoint;
                int n2 = currentElement.EndPoint;
                P1 = Nodes[n1 - 1].Loc;
                P2 = Nodes[n2 - 1].Loc;
                stop = PointOnLine(P, P1, P2);
                i++;
            }
            int revitID = Elements[i - 1].RevitID;
            Elements.RemoveAt(i - 1);
            int material = currentElement.MaterialNumber;
            int section = currentElement.SectionNumber;
            MidasElement element1 = new MidasElement(nElements, revitID, N, currentElement.EndPoint, section, material);
            nElements++;
            MidasElement element2 = new MidasElement(nElements, revitID, currentElement.StartPoint, N, section, material);
            nElements++;

            Elements.Add(element1);
            Elements.Add(element2);
        }

        public Result Execute(ExternalCommandData commandData, 
            ref string message,
            ElementSet elements)
        {
            //On accède aux objets les plus hauts

            UIApplication rvtUIApp = commandData.Application;
            UIDocument rvtUIDoc = rvtUIApp.ActiveUIDocument;
            m_rvtApp = rvtUIApp.Application;
            m_rvtDoc = rvtUIDoc.Document;

            FilteredElementCollector StructuralElements = GetStructuralElements(m_rvtDoc);
            FilteredElementCollector SupportAndLoads = GetSupportAndLoads(m_rvtDoc);


            List<MidasElement> Elements = new List<MidasElement>();
            List<MidasPoint> Nodes = new List<MidasPoint>();
            List<PointBoundary> Boundaries = new List<PointBoundary>();
            List<MidasLoadCase> LoadCases = new List<MidasLoadCase>();
            List<MidasPointLoad> PointLoads = new List<MidasPointLoad>();
            List<MidasElementLoad> ElementLoads = new List<MidasElementLoad>();
            List<MidasLineLoad> LineLoads = new List<MidasLineLoad>();
            List<string> MaterialList = new List<string>();
            List<string> SectionList = new List<string>();

            List<int> ElementNodes = new List<int>();

            int nElements = 1, nNodes = 1, nBoundaries = 1, nPointLoads = 1, nElementLoads = 1;

            
            var saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = ".mgt";
            saveFileDialog.ShowDialog();

            string FileToRead = "C:/API/template.mgt", FileToWrite = saveFileDialog.FileName;

            var txtLines = File.ReadAllLines(FileToRead).ToList();


            foreach (var elem in StructuralElements)
            {
                 getMidasParams(elem, ref nElements, ref nNodes, ref Nodes, ref Elements, ref MaterialList, ref SectionList, ref ElementNodes);
            }


            foreach (var elem in SupportAndLoads)
            {
                if (elem is BoundaryConditions)
                {
                    getBoundaryConditions(elem, ref nNodes, ref nBoundaries, ref Nodes, ref Boundaries);
                }
                else if (elem is PointLoad)
                {
                    getPointLoad(elem, ref nNodes, ref nPointLoads, ref Nodes, ref PointLoads, ref LoadCases);
                }
                else if (elem is LineLoad)
                {
                    getLineLoad(elem, ref nElementLoads, ref ElementLoads, ref LineLoads, ref LoadCases);
                }
            }

            int length = Elements.Count;

            foreach (var con in Boundaries)
            {
                int N = con.PointNumber;
                if (ElementNodes.IndexOf(N) == -1)
                {
                    splitElement(N, ref Elements, ref nElements, Nodes, length);
                }
            }

            foreach (var con in PointLoads)
            {
                int N = con.PointNumber;
                if (ElementNodes.IndexOf(N) == -1)
                {
                    splitElement(N, ref Elements, ref nElements, Nodes, length);
                }
            }

            foreach (var lineLoad in LineLoads)
            {
                LineToElementLoad(lineLoad, ref ElementLoads, ref nElementLoads, Nodes, Elements);
            }

            for (int n=0; n < MaterialList.Count(); n++)
            {
                string material = MaterialList[n];
                if (material[0] == 'S')
                {
                    string Material = (n + 1).ToString() + ", STEEL," + material + ", 0, 0, , C, NO, 0.02, 1, EN05(S)    ,            ," + material + ", NO, 2.1e+008 \n";
                    txtLines.Insert(txtLines.IndexOf("; EndOf Material"), Material);
                }
                else if (material[0] == 'C')
                {
                    string Material = (n + 1).ToString() + ", CONC," + material + ", 0, 0, , C, NO, 0.05, 1, EN04(RC)    ,            ," + material + ", NO, 3.2836e+007 \n";
                    txtLines.Insert(txtLines.IndexOf("; EndOf Material"), Material);
                }
            }

            Regex sectionRect = new Regex(@"([0-9]{2,4})\s+(x)\s+([0-9]{2,4})");
            Regex dim = new Regex(@"([0-9]{2,4})");

            for (int n = 0; n < SectionList.Count(); n++)
            {
                string section = SectionList[n];
                Match match = sectionRect.Match(section);
                if (match.Success)
                {
                    string dim1 = (Convert.ToInt32(match.Groups[1].Value)/1000.0).ToString().Replace(",",".");
                    string dim2 = (Convert.ToInt32(match.Groups[3].Value) / 1000.0).ToString().Replace(",", ".");
                    string Section = (n+1).ToString() + ", DBUSER, " + section + ", CC, 0, 0, 0, 0, 0, 0, YES, NO, SB , 2, " + dim2 + ", " + dim1 + ", 0, 0, 0, 0, 0, 0, 0, 0";
                    txtLines.Insert(txtLines.IndexOf("; EndOf Section"), Section);
                }
                else
                {
                    string Section = (n + 1).ToString() + ", DBUSER, " + SectionList[n] + ", CC, 0, 0, 0, 0, 0, 0, YES, NO, H  , 1, UNI," + SectionList[n];
                    txtLines.Insert(txtLines.IndexOf("; EndOf Section"), Section);
                }
            }

            txtLines.Insert(txtLines.IndexOf("; EndOf Nodes"), MidasGeometryString(Nodes));
            txtLines.Insert(txtLines.IndexOf("; EndOf Elements"), MidasElementString(Elements));
            txtLines.Insert(txtLines.IndexOf("; EndOf Support"), MidasBoundaryString(Boundaries));

            foreach (var line in MidasLoadCaseString(LoadCases))
            {
                txtLines.Insert(txtLines.IndexOf("; EndOf Load Cases"), line);
            }

            foreach (var pointLoad in PointLoads)
            {
                writePointLoadString(pointLoad, ref txtLines);
            }

            foreach (var elementLoad in ElementLoads)
            {
                writeElementLoadString(elementLoad, ref txtLines, Elements);
            }


            File.WriteAllLines(FileToWrite, txtLines);



            TaskDialog.Show("Revit - Envoi vers Midas", "Fichier .mgt écrit avec succès : \n\n" + FileToWrite + "\n\n Dans Midas, créer un nouveau fichier et cliquer sur l'icône Midas > Import > Midas GEN MGT File");

            //Process.Start("C:/Programmes/MIDAS/midas Gen.exe");

            return Result.Succeeded;
        }
    }

    
}
