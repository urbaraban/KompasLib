using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Tools;
using System;
using System.Windows;
using System.Windows.Media;
using reference = System.Int32;

namespace KompasLib.KompasTool
{
    public class ContourCalc
    {
        public static PathGeometry GetGeometry(KmpsDoc doc, double crs, bool add, bool cursor = true)
        {
            Point Offcet = new Point(0, 0);

            if (KmpsAppl.KompasAPI != null)
            {
                #region получение контура
                double x = 0, y = 0;
                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                //Ищем или находим макрообъект по индексу потолка
                IMacroObject macroObject = doc.Macro.FindCeilingMacro("0");
                //Создаем если нет
                if (macroObject == null) macroObject = doc.Macro.MakeCeilingMacro("0");

                if (cursor)
                {
                    if (!add)
                    {
                        doc.Macro.RemoveCeilingMacro("0");
                        macroObject = doc.Macro.MakeCeilingMacro("0");
                    }
                    doc.D5.ksCursor(info, ref x, ref y, 0);
                    if (!doc.Macro.AddCeilingMacro(doc.D5.ksMakeEncloseContours(0, x, y), "0")) MessageBox.Show("Контур не добавили", "Ошибка"); //Добавляем ksMakeEncloseContours
                }


                ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                if (doc.D5.ksGetObjGabaritRect(macroObject.Reference, spcGabarit) == 1)
                {
                    ksMathPointParam mathBop = spcGabarit.GetpBot();
                    ksMathPointParam mathTop = spcGabarit.GetpTop();
                    Offcet.X = mathBop.x + (mathTop.x - mathBop.x) / 2;
                    Offcet.Y = mathBop.y + (mathTop.y - mathBop.y) / 2;
                }

                ksInertiaParam inParam = (ksInertiaParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_InertiaParam);
                ksIterator Iterator1 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                Iterator1.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject.Reference);

                reference refContour1 = Iterator1.ksMoveIterator("F");
                PathGeometry path = new PathGeometry();

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
                        path.Figures.Add(TraceContour(contour));

                    refContour1 = Iterator1.ksMoveIterator("N"); //Двигаем итератор 1
                }

                Iterator1.ksDeleteIterator(); //Удаляем итератор 1 после полного перебора

                return path;

                #endregion
            }
            else MessageBox.Show("Объект не захвачен", "Сообщение");

            return null;

            PathFigure TraceContour(IContour contour)
            {
                if (contour != null)
                {
                    PathFigure pathFigure = new PathFigure();

                    for (int i = 0; i < contour.Count; i++)
                    {
                        IContourSegment pDrawObj = (IContourSegment)contour.Segment[i];
                        // Получить тип объекта

                        switch (pDrawObj.SegmentType)
                        {
                            case ksContourSegmentEnum.ksCSLineSeg:
                                IContourLineSegment contourLineSegment = (IContourLineSegment)pDrawObj;

                                if (i == 0)
                                    pathFigure.StartPoint = new Point(contourLineSegment.X1 - Offcet.X, -contourLineSegment.Y1 + Offcet.Y);

                                pathFigure.Segments.Add(new System.Windows.Media.LineSegment(new Point(contourLineSegment.X2 - Offcet.X, -contourLineSegment.Y2 + Offcet.Y), true));
                                break;
                            case ksContourSegmentEnum.ksCSArc:
                                IContourArc contourArc = (IContourArc)pDrawObj;

                                if (i == 0)
                                    pathFigure.StartPoint = new Point(contourArc.X1 - Offcet.X, -contourArc.Y1 + Offcet.Y);

                                double RotateAngel = contourArc.Direction ? Math.Abs(contourArc.Angle1 + 360 - contourArc.Angle2) % 360 : Math.Abs(contourArc.Angle2 + 360 - contourArc.Angle1) % 360;

                                pathFigure.Segments.Add(
                                                new ArcSegment(
                                                    new Point(contourArc.X2 - Offcet.X, -contourArc.Y2 + Offcet.Y),
                                                new Size(contourArc.Radius, contourArc.Radius), RotateAngel, RotateAngel > 180,
                                                contourArc.Direction ? SweepDirection.Clockwise : SweepDirection.Counterclockwise,
                                                true));

                                break;

                            default:
                                double[] arrayCurve = pDrawObj.Curve2D.CalculatePolygonByStep(crs / pDrawObj.Curve2D.Length);
                                PolyLineSegment polyLineSegment = new PolyLineSegment();

                                if (i == 0)
                                    pathFigure.StartPoint = new Point(arrayCurve[0] - Offcet.X, -arrayCurve[1] + Offcet.Y);

                                for (int j = 2; j < arrayCurve.Length; j += 2)
                                {
                                    polyLineSegment.Points.Add(new Point(arrayCurve[j] - Offcet.X, -arrayCurve[j + 1] + Offcet.Y));
                                }
                                pathFigure.Segments.Add(polyLineSegment);
                                break;
                        }
                    }
                    pathFigure.IsClosed = true;
                    return pathFigure;
                }
                return null;
            }
        }

    }
}
