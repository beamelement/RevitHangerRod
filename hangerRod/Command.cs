using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace hangerRod
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document revitDoc = commandData.Application.ActiveUIDocument.Document;  //取得文档           
            Application revitApp = commandData.Application.Application;             //取得应用程序            
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档        


            Window1 window1 = new Window1();
            Color colorOrigin = new Color(0, 0, 0);
            FamilyInstance familyInstance = null;


            //窗口输入参数模块
            if (window1.ShowDialog() == true)
            {
                //窗口打开并停留，只有点击按键之后，窗口关闭并返回true
            }

            //按键会改变window的属性，通过对属性的循环判断来实现对按键的监测
            while (!window1.Done)
            {

                //选择一个腹杆
                if (window1.Selected)
                {
                    Selection sel = uiDoc.Selection;
                    Reference ref1 = sel.PickObject(ObjectType.Element, "选择一个横梁");
                    Element elem = revitDoc.GetElement(ref1);
                    familyInstance = elem as FamilyInstance;

                    using (Transaction tran = new Transaction(uiDoc.Document))
                    {
                        tran.Start("选中横梁");

                        //初试颜色
                        Material materialOrigin = uiDoc.Document.GetElement(familyInstance.GetMaterialIds(false).First()) as Material;
                        colorOrigin = materialOrigin.Color;

                        //改变选中实例的线的颜色
                        OverrideGraphicSettings overrideGraphicSettings = new OverrideGraphicSettings();
                        overrideGraphicSettings = uiDoc.Document.ActiveView.GetElementOverrides(familyInstance.Id);

                        Color color = new Color(255, 0, 0);
                        overrideGraphicSettings.SetProjectionLineColor(color);
                        //在当前视图下设置，其它视图保持原来的
                        uiDoc.Document.ActiveView.SetElementOverrides(familyInstance.Id, overrideGraphicSettings);
                        uiDoc.Document.Regenerate();



                        //获取吊杆底部坐标
                        XYZ instancePoint = (familyInstance.Location as LocationPoint).Point;//实例位置
                        XYZ instanceNormalVec = familyInstance.FacingOrientation; //实例的法向量
                        XYZ instanceTangentVec = instanceNormalVec.CrossProduct(XYZ.BasisZ);//实例切向量
                        double tangentDistance = familyInstance.LookupParameter("l1/2").AsDouble() + familyInstance.LookupParameter("l2").AsDouble() + familyInstance.LookupParameter("l3").AsDouble() / 2; //获取吊杆圆心坐标
                        double ZDistance = 10;//竖向的偏移距离，之后有时间可以考虑把他设坐参数
                        XYZ ptemp = (instancePoint + instanceTangentVec * tangentDistance);
                        XYZ hangerRodDownPoint = new XYZ(ptemp.X, ptemp.Y, ptemp.Z - ZDistance);//横梁底部坐标



                        //获取吊杆顶部坐标
                        View3D view3d = uiDoc.Document.ActiveView as View3D;
                        List<FamilyInstance> instances = ReferenceIntersectElement(uiDoc.Document, view3d, hangerRodDownPoint, XYZ.BasisZ);//射线法，找元素
                        FamilyInstance chordInstance = instances.Last();
                        XYZ hangerRodUPPoint = Interpolation(chordInstance, hangerRodDownPoint, commandData);//内插法求得插入的点

                        //创建模型线
                        Line hangerRodLine = Line.CreateBound(hangerRodUPPoint, hangerRodDownPoint);
                        XYZ hangerRodVec = new XYZ(hangerRodUPPoint.X - hangerRodDownPoint.X, hangerRodUPPoint.Y - hangerRodDownPoint.Y, hangerRodUPPoint.Z - hangerRodDownPoint.Z);
                        XYZ vec = new XYZ(1, 1, 1);
                        XYZ normal = hangerRodVec.CrossProduct(vec);
                        Plane plane = Plane.CreateByNormalAndOrigin(normal, hangerRodUPPoint);
                        SketchPlane sketchPlane = SketchPlane.Create(uiDoc.Document, plane);
                        ModelLine modelLine = uiDoc.Document.Create.NewModelCurve(hangerRodLine, sketchPlane) as ModelLine;



                        tran.Commit();
                    }

                    window1.Selected = false;
                }


                if (window1.ShowDialog() == true)
                {
                    //窗口打开并停留，只有点击按键之后，窗口关闭并返回true
                }
            }


            //恢复轮廓线颜色
            using (Transaction tran = new Transaction(uiDoc.Document))
            {
                tran.Start("恢复轮廓线的颜色");

                //恢复标注前的颜色

                OverrideGraphicSettings overrideGraphicSettings = new OverrideGraphicSettings();
                overrideGraphicSettings = uiDoc.Document.ActiveView.GetElementOverrides(familyInstance.Id);
                overrideGraphicSettings.SetProjectionLineColor(colorOrigin);
                //在当前视图下设置，其它视图保持原来的
                uiDoc.Document.ActiveView.SetElementOverrides(familyInstance.Id, overrideGraphicSettings);
                uiDoc.Document.Regenerate();

                tran.Commit();
            }



            //载入族
            FamilySymbol familySymbol;
            using (Transaction tran = new Transaction(uiDoc.Document))
            {
                tran.Start("载入族");

                //载入弦杆族
                string file = @"C:\Users\zyx\Desktop\2RevitArcBridge\RevitWindBrace\RevitWindBrace\Source\WindBrace.rfa";
                familySymbol = loadFaimly(file, commandData);
                familySymbol.Activate();

                tran.Commit();
            }

            ////风撑族实例化
            //using (Transaction tran = new Transaction(uiDoc.Document))
            //{
            //    tran.Start("创建风撑");

            //    createWindBrace(PPosition, familySymbol, commandData);

            //    tran.Commit();
            //}




            return Result.Succeeded;

        }








        //点的内插
        private XYZ Interpolation(FamilyInstance chordInstance, XYZ hangerRodDownPoint, ExternalCommandData commandData)
        {
            Document revitDoc = commandData.Application.ActiveUIDocument.Document;  //取得文档           
            Application revitApp = commandData.Application.Application;             //取得应用程序            
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档  

            IList<ElementId> chordPointIDs = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(chordInstance);
            List<XYZ> chordPoints = new List<XYZ>();
            foreach (ElementId elementId in chordPointIDs)
            {
                Element e = uiDoc.Document.GetElement(elementId);
                ReferencePoint referencePoint = e as ReferencePoint;
                chordPoints.Add(referencePoint.Position);
            }

            XYZ p1 = new XYZ(chordPoints[0].X, chordPoints[0].Y, 0);
            XYZ p2 = new XYZ(chordPoints[1].X, chordPoints[1].Y, 0);
            XYZ p3 = new XYZ(hangerRodDownPoint.X, hangerRodDownPoint.Y, 0);


            Line line = Line.CreateBound(chordPoints[0], chordPoints[1]);//三维线
            Line line1 = Line.CreateBound(p1, p2);//投影线
            double L = Line.CreateBound(p1, p3).Length;//到起点的距离
            double ratio = L / line1.Length;
            XYZ UPPoint = line.Evaluate(ratio, true);
            return UPPoint;
        }


        //载入族
        private FamilySymbol loadFaimly(string file, ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档     
            bool loadSuccess = uiDoc.Document.LoadFamily(file, out Family family);

            if (loadSuccess)
            {
                //假如成功导入
                //得到族模板
                ElementId elementId;
                ISet<ElementId> symbols = family.GetFamilySymbolIds();
                elementId = symbols.First();
                FamilySymbol adaptiveFamilySymbol = uiDoc.Document.GetElement(elementId) as FamilySymbol;

                return adaptiveFamilySymbol;
            }
            else
            {
                //假如已经导入,则通过名字找到这个族
                FilteredElementCollector collector = new FilteredElementCollector(uiDoc.Document);
                collector.OfClass(typeof(Family));//过滤得到文档中所有的族
                IList<Element> families = collector.ToElements();
                FamilySymbol adaptiveFamilySymbol = null;
                foreach (Element e in families)
                {

                    Family f = e as Family;
                    //通过名字进行筛选
                    if (f.Name == "WindBrace")
                    {
                        adaptiveFamilySymbol = uiDoc.Document.GetElement(f.GetFamilySymbolIds().First()) as FamilySymbol;
                    }
                }
                return adaptiveFamilySymbol;

            }
        }


        //风箱族实例化
        private void createWindBrace(List<XYZ> points, FamilySymbol FamilySymbol, ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档    

            //创建实例，并获取其自适应点列表
            FamilyInstance familyInstance = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(uiDoc.Document, FamilySymbol);
            IList<ElementId> adaptivePoints = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(familyInstance);

            for (int i = 0; i < points.Count; i += 1)
            {
                //取得的参照点
                ReferencePoint referencePoint = uiDoc.Document.GetElement(adaptivePoints[i]) as ReferencePoint;
                //设置参照点坐标
                referencePoint.Position = points[i];
            }
        }


        //射线法找元素
        private List<FamilyInstance> ReferenceIntersectElement(Document doc, View3D view3d, XYZ origin, XYZ normal)
        {

            ElementClassFilter filter = new ElementClassFilter(typeof(FamilyInstance));
            ReferenceIntersector refInter = new ReferenceIntersector(filter, FindReferenceTarget.Element, view3d);
            IList<ReferenceWithContext> listContext = refInter.Find(origin, normal);

            List<FamilyInstance> instances = new List<FamilyInstance>();

            foreach (ReferenceWithContext reference in listContext)
            {
                Reference refer = reference.GetReference();
                ElementId id = refer.ElementId;
                FamilyInstance instance = doc.GetElement(id) as FamilyInstance;

                if (instance.Symbol.Family.Name.Contains("chordFamlily"))
                {

                    instances.Add(instance);

                }
            }

            return instances;
        }

    }
}
