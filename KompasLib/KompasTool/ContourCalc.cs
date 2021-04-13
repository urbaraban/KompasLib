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
        public static GeometryGroup GetGeometry(KmpsDoc doc, double crs, bool add, bool cursor = true)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IMacroObject macroObject = doc.Macro.FindCeilingMacro("0");
                #region получение контура
                if (doc.GetSelectContainer().SelectedObjects != null)
                {
                    macroObject = SelectObjToMacro(doc.GetSelectContainer());
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
                        foreach (Geometry geometry in TraceContour(contour).Children)
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

            GeometryGroup TraceContour(IContour contour)
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
                                double[] arrayCurve = pDrawObj.Curve2D.CalculatePolygonByStep(crs / pDrawObj.Curve2D.Length);
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

            IMacroObject SelectObjToMacro(ISelectionManager selectionManager)
            {
                List<IContour> contours = GetContourFromSelectObj(selectionManager);

                doc.Macro.RemoveCeilingMacro("0");
                doc.Macro.MakeCeilingMacro("0");

                foreach (IContour contour in contours)
                {
                    doc.Macro.AddCeilingMacro(contour.Reference, "0");
                }

                return doc.Macro.FindCeilingMacro("0");
            }

            List<IContour> GetContourFromSelectObj(ISelectionManager selectionManager)
            {
                if (selectionManager.SelectedObjects != null)
                {
                    List<IContour> contours = new List<IContour>();

                    int[] YesType = { 1, 2, 3, 8, 26, 28, 31, 32, 33, 34, 35, 36, 80 };

                    IDrawingContours drawingContours = doc.GetDrawingContainer().DrawingContours;

                    // Получить массив объектов
                    try
                    {
                        Array arrS = (Array)selectionManager.SelectedObjects;

                        foreach (object obj in arrS) contours.Add(MakeContour(obj));
                    }
                    catch
                    {
                        //если один объект
                        contours.Add(MakeContour(selectionManager.SelectedObjects));
                    }
                    

                    return contours;

                    IContour MakeContour(object obj)
                    {
                        IDrawingContour drawingContour = drawingContours.Add();
                        IContour contour = (IContour)drawingContour;
                        IDrawingObject pObj = (IDrawingObject)obj;

                        if (IndexOfTrue(YesType, (int)pObj.DrawingObjectType))
                            contour.CopyCurve(pObj, false);
                        drawingContour.Update();
                        return contour;
                    }
                }
                return null;

                bool IndexOfTrue(int[] arr, int value)
                {
                    for (int i = 0; i < arr.Length; i++)
                        if (arr[i] == value) return true;
                    return false;
                }

            }
        }
    }
}
