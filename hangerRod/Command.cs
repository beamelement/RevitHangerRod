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

                //选择横梁
                if (window1.Selected)
                {
                    Selection sel = uiDoc.Selection;
                    Reference ref1 = sel.PickObject(ObjectType.Element, "选择一个横梁");
                    Element elem = revitDoc.GetElement(ref1);
                    familyInstance = elem as FamilyInstance;

                    using (Transaction tran = new Transaction(uiDoc.Document))
                    {
                        tran.Start("创建吊杆模型线");

                        //改变选中实例轮廓颜色
                        colorOrigin = changeProfileColor(familyInstance,commandData);

                        //创建吊杆所要用到的模型线
                        List<ModelLine>modelLines = createHangerROdModelLines(familyInstance,commandData);
                        //注意此处模型线底部高度是随意设置的，到时候要进行修改

                        FamilySymbol familySymbol;
                        //载入吊杆族
                        string file = @"C:\Users\zyx\Desktop\2RevitArcBridge\hangerRod\hangerRod\source\hangerRod.rfa";
                        familySymbol = loadFaimly(file, commandData);
                        familySymbol.Activate();

                        ////创建族实例
                        //createHangerRod(modelLine, familySymbol, commandData); 创建基于线的族实例失效，但是已经定义好工作平面，此步可以通过手动实现


                        //创建并激活形成一个工作平面
                        XYZ p1 = modelLines[0].GeometryCurve.Evaluate(0, true);
                        XYZ p2 = modelLines[0].GeometryCurve.Evaluate(1, true);
                        XYZ p3 = modelLines[1].GeometryCurve.Evaluate(0, true);
                        XYZ p4 = modelLines[1].GeometryCurve.Evaluate(1, true);

                        XYZ vec1 = p1.Subtract(p4);
                        XYZ vec2 = p2.Subtract(p3);
                        XYZ normal = vec1.CrossProduct(vec2);

                        Plane plane = Plane.CreateByNormalAndOrigin(normal, p1);
                        SketchPlane sketchPlane = SketchPlane.Create(uiDoc.Document, plane);


                        uiDoc.Document.ActiveView.SketchPlane = sketchPlane;
                        uiDoc.Document.ActiveView.ShowActiveWorkPlane();
                        Reference reference = sketchPlane.GetPlaneReference();

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
                
                //恢复实例轮廓颜色
                recoverProfileColor(familyInstance, colorOrigin, commandData);

                tran.Commit();
            }


            return Result.Succeeded;

        }







        private List<ModelLine> createHangerROdModelLines(FamilyInstance familyInstance,ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档  

            List<ModelLine> modelLines = new List<ModelLine>();

            for(int i = 0; i < 3; i += 2)
            {
                //获取吊杆底部坐标
                XYZ instancePoint = (familyInstance.Location as LocationPoint).Point;//实例位置
                XYZ instanceNormalVec = familyInstance.FacingOrientation; //实例的法向量
                XYZ instanceTangentVec = instanceNormalVec.CrossProduct(XYZ.BasisZ);//实例切向量
                double tempDistance = familyInstance.LookupParameter("l1/2").AsDouble() + familyInstance.LookupParameter("l2").AsDouble() + familyInstance.LookupParameter("l3").AsDouble() / 2; //获取吊杆圆心坐标
                double distance = (i - 1) * tempDistance;
                double ZDistance = 10;//竖向的偏移距离，之后有时间可以考虑把他设坐参数
                XYZ ptemp = (instancePoint + instanceTangentVec * distance);
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
                
                modelLines.Add(modelLine);
            }


            return modelLines;


        }











        //恢复选中实例的颜色
        private void recoverProfileColor(FamilyInstance familyInstance, Color colorOrigin, ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档  
            //恢复标注前的颜色
            OverrideGraphicSettings overrideGraphicSettings = new OverrideGraphicSettings();
            overrideGraphicSettings = uiDoc.Document.ActiveView.GetElementOverrides(familyInstance.Id);
            overrideGraphicSettings.SetProjectionLineColor(colorOrigin);
            //在当前视图下设置，其它视图保持原来的
            uiDoc.Document.ActiveView.SetElementOverrides(familyInstance.Id, overrideGraphicSettings);
            uiDoc.Document.Regenerate();
        }

        //改变选中实例的颜色(被选中的族实例要已经赋予了材质，这个方法才起作用)
        private Color changeProfileColor(FamilyInstance familyInstance,ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档  

            //初试颜色
            Material materialOrigin = uiDoc.Document.GetElement(familyInstance.GetMaterialIds(false).First()) as Material;
            Color colorOrigin = materialOrigin.Color;

            //改变选中实例的线的颜色
            OverrideGraphicSettings overrideGraphicSettings = new OverrideGraphicSettings();
            overrideGraphicSettings = uiDoc.Document.ActiveView.GetElementOverrides(familyInstance.Id);

            Color color = new Color(255, 0, 0);
            overrideGraphicSettings.SetProjectionLineColor(color);
            //在当前视图下设置，其它视图保持原来的
            uiDoc.Document.ActiveView.SetElementOverrides(familyInstance.Id, overrideGraphicSettings);
            uiDoc.Document.Regenerate();

            return colorOrigin;

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
                    if (f.Name == "hangerRod")
                    {
                        adaptiveFamilySymbol = uiDoc.Document.GetElement(f.GetFamilySymbolIds().First()) as FamilySymbol;
                    }
                }
                return adaptiveFamilySymbol;

            }
        }



        //吊杆族实例化
        private void createHangerRod(ModelLine modelLine, FamilySymbol familySymbol, ExternalCommandData commandData)
        {
            Document revitDoc = commandData.Application.ActiveUIDocument.Document;  //取得文档           
            Application revitApp = commandData.Application.Application;             //取得应用程序            
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;           //取得当前活动文档      
            //依据线创建一个包含这条线的面
            XYZ p1 = modelLine.GeometryCurve.GetEndPoint(0);
            XYZ p2 = modelLine.GeometryCurve.GetEndPoint(1);
            XYZ LineVec = new XYZ(p1.X - p2.X, p1.Y - p2.Y, p1.Z - p2.Z);
            XYZ vectemp = new XYZ(1, 1, 1);
            XYZ normal = LineVec.CrossProduct(vectemp);
            Plane plane = Plane.CreateByNormalAndOrigin(normal, p1);
            SketchPlane sketchPlane = SketchPlane.Create(uiDoc.Document, plane);
            Reference reference = sketchPlane.GetPlaneReference();





            //Line line = modelLine.GeometryCurve as Line;


            Line line = Line.CreateBound(p1, p2);

            FamilyInstance familyInstance = revitDoc.Create.NewFamilyInstance(reference, line, familySymbol);
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
