using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Tools;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using reference = System.Int32;

namespace KompasLib.KompasTool
{
    public class ContourCalc
    {
        public static int[] YesTypeObject = { 1, 2, 3, 8, 26, 28, 31, 32, 33, 34, 35, 36, 80 };

        public static GeometryGroup GetGeometry(KmpsDoc doc, double crs, bool add, bool cursor = true)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IMacroObject macroObject = doc.Macro.FindCeilingMacro("0");
                #region получение контура
                if (doc.GetSelectContainer().SelectedObjects != null)
                {
                    return GetSelectGeometry(doc.GetSelectContainer(), doc.GetDrawingContainer().DrawingContours, crs);
                }
                else
                {
                    double x = 0, y = 0;
                    RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                    //Ищем или находим макрообъект по индексу потолка
                    
                    //Создаем если нет
                    if (macroObject == null) macroObject = doc.Macro.MakeCeilingMacro("0");

                    if (cursor)
                    {
                        if (add == false)
                        {
                            doc.Macro.RemoveCeilingMacro("0");
                            macroObject = doc.Macro.MakeCeilingMacro("0");
                        }
                        doc.D5.ksCursor(info, ref x, ref y, 0);
                        if (!doc.Macro.AddCeilingMacro(doc.D5.ksMakeEncloseContours(0, x, y), "0")) MessageBox.Show("Контур не добавили", "Ошибка"); //Добавляем ksMakeEncloseContours
                    }
                }


                ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                if (doc.D5.ksGetObjGabaritRect(macroObject.Reference, spcGabarit) == 1)
                {
                    ksMathPointParam mathBop = spcGabarit.GetpBot();
                    ksMathPointParam mathTop = spcGabarit.GetpTop();

                }

                ksInertiaParam inParam = (ksInertiaParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_InertiaParam);
                ksIterator Iterator1 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                Iterator1.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject.Reference);

                reference refContour1 = Iterator1.ksMoveIterator("F");

                GeometryGroup geometryGroup = new GeometryGroup();

                // contoursList.DisplayName = doc.D7.Name;

                //
                //Начинаем перебор контуров со всем что есть
                //
                //Заходим в первый контур
                refContour1 = Iterator1.ksMoveIterator("F");

                while (refContour1 != 0)
                {
                    IContour contour = doc.Macro.GiveContour(refContour1);

                    if (contour != null) 
                    {
                        foreach (Geometry geometry in TraceContour(contour, crs).Children)
                        {
                            geometryGroup.Children.Add(geometry);
                        }
                    }

                    refContour1 = Iterator1.ksMoveIterator("N"); //Двигаем итератор 1
                }

                Iterator1.ksDeleteIterator(); //Удаляем итератор 1 после полного перебора

                return geometryGroup;

                #endregion
            }
            else MessageBox.Show("Объект не захвачен", "Сообщение");

            return null;

           
            
        }


        public static List<IContour> GetContourFromSelectObj(ISelectionManager selectionManager, KmpsDoc doc)
        {
            if (selectionManager.SelectedObjects != null)
            {
                List<IContour> contours = new List<IContour>();

                IDrawingContours drawingContours = doc.GetDrawingContainer().DrawingContours;

                // Получить массив объектов
                if (selectionManager.SelectedObjects is Array array)
                {
                    foreach (object obj in array) contours.Add(MakeContour(obj, drawingContours));
                }
                else
                {
                    //если один объект
                    contours.Add(MakeContour(selectionManager.SelectedObjects, drawingContours));
                }


                return contours;
            }
            return null;
        }

        public static IContour MakeContour(object obj, IDrawingContours drawingContours)
        {
            IDrawingContour drawingContour = drawingContours.Add();
            IContour contour = (IContour)drawingContour;
            IDrawingObject pObj = (IDrawingObject)obj;

            if (IndexOfTrue(YesTypeObject, (int)pObj.DrawingObjectType))
                contour.CopyCurve(pObj, false);
            drawingContour.Update();
            return contour;
        }

        public static GeometryGroup TraceContour(IContour contour, double CRS)
        {
            GeometryGroup geometryGroup = new GeometryGroup();
            if (contour != null)
            {
                for (int i = 0; i < contour.Count; i++)
                {
                    IContourSegment pDrawObj = (IContourSegment)contour.Segment[i];
                    // Получить тип объекта

                    switch (pDrawObj.SegmentType)
                    {
                        case ksContourSegmentEnum.ksCSLineSeg:
                            IContourLineSegment contourLineSegment = (IContourLineSegment)pDrawObj;
                            geometryGroup.Children.Add(
                                new LineGeometry(new Point(contourLineSegment.X1, -contourLineSegment.Y1),
                                new Point(contourLineSegment.X2, -contourLineSegment.Y2)));
                            break;
                        case ksContourSegmentEnum.ksCSArc:
                            IContourArc contourArc = (IContourArc)pDrawObj;

                            PathFigure pathFigure = new PathFigure();

                            pathFigure.StartPoint = new Point(contourArc.X1, -contourArc.Y1);

                            double RotateAngel = contourArc.Direction ? Math.Abs(contourArc.Angle1 + 360 - contourArc.Angle2) % 360 : Math.Abs(contourArc.Angle2 + 360 - contourArc.Angle1) % 360;

                            pathFigure.Segments.Add(
                                            new ArcSegment(
                                                new Point(contourArc.X2, -contourArc.Y2),
                                            new Size(contourArc.Radius, contourArc.Radius), RotateAngel, RotateAngel > 180,
                                            contourArc.Direction ? SweepDirection.Clockwise : SweepDirection.Counterclockwise,
                                            true));

                            geometryGroup.Children.Add(new PathGeometry(new List<PathFigure>() { pathFigure }));
                            break;

                        default:
                            double[] arrayCurve = pDrawObj.Curve2D.CalculatePolygonByStep(CRS / pDrawObj.Curve2D.Length);
                            PolyLineSegment polyLineSegment = new PolyLineSegment();

                            PathFigure pathFigureCurve = new PathFigure();

                            pathFigureCurve.StartPoint = new Point(arrayCurve[0], -arrayCurve[1]);

                            for (int j = 2; j < arrayCurve.Length; j += 2)
                            {
                                polyLineSegment.Points.Add(new Point(arrayCurve[j], -arrayCurve[j + 1]));
                            }
                            pathFigureCurve.Segments.Add(polyLineSegment);

                            geometryGroup.Children.Add(new PathGeometry(new List<PathFigure>() { pathFigureCurve }));
                            break;
                    }
                }
                //pathFigure.IsClosed = contour.Closed;

                return geometryGroup;
            }
            return null;
        }

        public static GeometryGroup GetSelectGeometry(ISelectionManager selectionManager, IDrawingContours drawingContours, double CRS)
        {
            if (selectionManager.SelectedObjects != null)
            {
                GeometryGroup geometryGroup = new GeometryGroup();

                // Получить массив объектов
                if (selectionManager.SelectedObjects is Array array)
                {
                    foreach (object obj in array)
                    {
                        CheckAddObject(GetObjGeometry(obj, CRS, drawingContours), geometryGroup);
                    }
                }
                else
                {
                    //если один объект
                    CheckAddObject(GetObjGeometry(selectionManager.SelectedObjects, CRS, drawingContours), geometryGroup);
                }

                return geometryGroup;
            }
            return null;

            void CheckAddObject(Geometry geometry, GeometryGroup geometryGroup)
            {
                if (geometry != null) geometryGroup.Children.Add(geometry);
            }
        }

        public static Geometry GetObjGeometry(object obj, double CRS, IDrawingContours drawingContours)
        {
            if (obj is IDrawingObject drawingObject)
            {
                switch (drawingObject.DrawingObjectType)
                {
                    case DrawingObjectTypeEnum.ksDrPoint:
                        IPoint point = (IPoint)drawingObject;
                        return null;

                    case DrawingObjectTypeEnum.ksDrLineSeg:
                        ILineSegment lineSegment = (ILineSegment)drawingObject;
                        return new LineGeometry(new Point(lineSegment.X1, -lineSegment.Y1),
                                                new Point(lineSegment.X2, -lineSegment.Y2));

                    case DrawingObjectTypeEnum.ksDrArc:
                        IArc arc = (IArc)drawingObject;

                        PathFigure pathFigure = new PathFigure();

                        pathFigure.StartPoint = new Point(arc.X1, -arc.Y1);

                        double RotateAngel = arc.Direction ? Math.Abs(arc.Angle1 + 360 - arc.Angle2) % 360 : Math.Abs(arc.Angle2 + 360 - arc.Angle1) % 360;

                        pathFigure.Segments.Add(
                                        new ArcSegment(
                                            new Point(arc.X2, -arc.Y2),
                                        new Size(arc.Radius, arc.Radius), RotateAngel, RotateAngel > 180,
                                        arc.Direction ? SweepDirection.Clockwise : SweepDirection.Counterclockwise,
                                        true));

                        return new PathGeometry(new List<PathFigure>() { pathFigure });
                    default:
                        return TraceContour(MakeContour(drawingObject, drawingContours), CRS);
                }
            }
            return null;
        }

        public static bool IndexOfTrue(int[] arr, int value)
        {
            for (int i = 0; i < arr.Length; i++)
                if (arr[i] == value) return true;
            return false;
        }
    }
}
