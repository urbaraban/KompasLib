using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using System;
using System.Collections.Generic;
using System.Windows;
using reference = System.Int32;

namespace KompasLib.Tools
{
    public class SizeTool
    {
        private KmpsAppl api;

        public SizeTool(KmpsDoc DOC, KmpsAppl APPL)
        {
            this.api = APPL;
        }

        public List<ComboData> GetLineDimentionList()
        {
            ILineDimensions lineDimension = (ILineDimensions)api.Doc.GetSymbols2DContainer().LineDimensions;


            List<ComboData> _lineDimensions = new List<ComboData>();

            for (int i = 0; i < lineDimension.Count; i++)
            {
                ILineDimension dimension = lineDimension.LineDimension[i];
                IDimensionText dimensionText = (IDimensionText)dimension;
                if (dimensionText.Brackets == ksDimensionTextBracketsEnum.ksDimBracketsOff)
                {
                    IDrawingObject1 object1 = (IDrawingObject1)dimension;
                    if ((object1 != null) && (object1.Constraints != null))
                        foreach (IParametriticConstraint constraint in object1.Constraints) //перебираем ограничение
                            if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable) //и добавляем только с параметром
                                _lineDimensions.Add(new ComboData("Размер " + i, dimension.Reference));
                }
            }

            return _lineDimensions;
        }

        public class ComboData
        {

            public string Name { get; set; }
            public Int32 Reference { get; set; }

            public ComboData(string _name, Int32 _ref)
            {
                Name = _name;
                Reference = _ref;
            }
        }

        public void Coordinate(string index, string connectionstr, double CoordDopusk)
        {
            ILayer layer = api.Doc.GiveLayer(88);
            List<Point> points = new List<Point>();
            SQLTool sQL = new SQLTool();
            if (api.Doc != null)
            {
                ISelectionManager selection = api.Doc.GetSelectContainer();
                if (selection.SelectedObjects != null)
                {
                    IDrawingGroup TempGroup = (IDrawingGroup)KmpsAppl.KompasAPI.TransferReference(api.Doc.D5.ksNewGroup(0), api.Doc.D5.reference);
                    try
                    {
                        int[] YesType = { 1, 2, 3, 8, 26, 28, 31, 32, 33, 34, 35, 36, 80 };
                        int[] DimType = { 9, 10, 13, 14, 15, 43};
                        Array arrS = (Array)selection.SelectedObjects;
                        int i = 0;
                        api.ProgressBar.Start(i, arrS.Length, "Скрываем размеры", true);

                        bool tempflag = MessageBox.Show("Скрыть размеры?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes ? true : false;
                        foreach (object obj in arrS)
                        {
                            IDrawingObject pObj = (IDrawingObject)obj;
                            if (IndexOfTrue(YesType, (int)pObj.DrawingObjectType))
                                TempGroup.AddObjects(obj);
                            else if (IndexOfTrue(DimType, (int)pObj.DrawingObjectType)) 
                            {  pObj.LayerNumber = layer.LayerNumber; pObj.Update(); }
                            api.ProgressBar.SetProgress(i++, "Скрываем размеры", true);
                        }
                        api.ProgressBar.Stop("Закончили", true);
                    }
                    catch
                    {
                        object pObj = selection.SelectedObjects;
                        TempGroup.AddObjects(pObj);
                    }
                    api.Doc.VisibleLayer(77, true);
                    api.Doc.VisibleLayer(88, false);

                    TempGroup.Close();

                    ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                    ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                    if (api.Doc.D5.ksGetObjGabaritRect(TempGroup.Reference, spcGabarit) == 1)
                    {
                        ksMathPointParam mathBop = spcGabarit.GetpBot();
                        ksMathPointParam mathTop = spcGabarit.GetpTop();
                        //Координаты области
                        double x = mathBop.x;
                        double y = mathBop.y;
                        double dx = mathTop.x;
                        double dy = mathTop.y;
                        //Ее размер
                        double sizeX = Math.Round(Math.Abs(x - dx), 2);
                        double sizeY = Math.Round(Math.Abs(y - dy), 2);


                        //Запрос на ширину материала
                        double width = 320;
                        if (api.Doc.Variable.Give("factura", index) > -1)
                            width = double.Parse(sQL.ReturnValue("SELECT TOP 1 width FROM dbo.Factura WHERE IDFactura=" + api.Doc.Variable.Give("factura", index), "width", connectionstr));


                        // Выясняем где ширина выгоднее
                        bool xflag = true;
                        if ((sizeX / width) <= 1 && (sizeX / width) > (sizeY / width) || (sizeY > width))
                        {
                            // mathTop.x = x + width;
                            xflag = false;
                        }
                        else
                        {
                            // mathTop.y = y + width;
                            xflag = true;
                        }
                        //Получаем новые координаты области
                        dx = mathTop.x;
                        dy = mathTop.y;

                        double kooff = width * 0.1;

                        api.Doc.D5.ksEndObj();

                        ViewsAndLayersManager ViewsMng = api.Doc.D7.ViewsAndLayersManager;
                        IViews views = ViewsMng.Views;
                        IView view = views.ActiveView;

                        layer = api.Doc.GiveLayer(77);

                        layer.Color = 8355711;
                        layer.Update();


                        if (api.Doc.D5.ksMacro(0) == 1)
                        {
                            //Вертикальная
                            api.Doc.D5.ksLineSeg(x, y - kooff, x, dy + kooff, 3);
                            api.Doc.D5.ksLineSeg(dx, y - kooff, dx, dy + kooff, 3);
                            //Горизонтальные
                            api.Doc.D5.ksLineSeg(x - kooff, y, dx + kooff, y, 3);
                            api.Doc.D5.ksLineSeg(x - kooff, dy, dx + kooff, dy, 3);
                            //Центр
                            api.Doc.D5.ksLineSeg(x + (dx - x) / 2, y - kooff, x + (dx - x) / 2, dy + kooff, 3);
                            api.Doc.D5.ksLineSeg(x - kooff, y + (dy - y) / 2, dx + kooff, y + (dy - y) / 2, 3);
                            //Ширина
                            SetText(Math.Round(mathTop.x - mathBop.x, 1).ToString(), mathBop.x + (mathTop.x - mathBop.x) / 2, mathBop.y - kooff * 1.5, ksAllocationEnum.ksAlCentre, width / 40 * 2, 0, true);
                            //Высота
                            SetText(Math.Round(mathTop.y - mathBop.y, 1).ToString(), mathBop.x - kooff, mathBop.y + (mathTop.y - mathBop.y) / 2, ksAllocationEnum.ksAlCentre, width / 40 * 2, 90, true);

                            //Подпись Y
                            SetText((!xflag ? "<<" : string.Empty) + "Y", x, y - kooff, ksAllocationEnum.ksAlLeft, width / 40 * 1.5, -90);
                            //Подпись Х
                            SetText("X" + (xflag ? ">>" : string.Empty), x - kooff, y, ksAllocationEnum.ksAlRight, width / 40 * 1.5);
                            //Угол А
                            SetText("А", x - kooff / 4, y - kooff / 4, ksAllocationEnum.ksAlRight, width / 40 * 1.5);

                            ///
                            /// Идем по объектам и проставляем подписи
                            ///

                            try
                            {
                                Array arrS = (Array)TempGroup.Objects[0];
                                api.ProgressBar.Start(0, arrS.Length, "Точка:", true);
                                for (int i = 0; i < arrS.Length; i++)
                                {
                                    api.ProgressBar.SetProgress(i, "Точка:", true);
                                    SetPointToObj(arrS.GetValue(i), mathBop, mathTop, width);
                                }
                                api.ProgressBar.Stop("Закончили", true);
                            }
                            catch
                            {
                                SetPointToObj(TempGroup.Objects[0], mathBop, mathTop, width);
                            }

                            //конец макрообъекта
                            reference macroRef = api.Doc.D5.ksEndObj();
                            IMacroObject pMacroObj = (IMacroObject)KmpsAppl.KompasAPI.TransferReference(macroRef, api.Doc.D5.reference);

                            pMacroObj.Name = "LineMacro";
                            pMacroObj.LayerNumber = 77;
                            pMacroObj.Update();

                           
                            /*if (attribute != null)
                                {
                                    attribute.AddRow(string.Empty, 0);
                                    attribute.SetValue(string.Empty, 0, 0, );
                                    attribute.SetValue(string.Empty, 0, 1, mathBop.y);
                                    attribute.SetValue(string.Empty, 0, 2, mathTop.x);
                                    attribute.SetValue(string.Empty, 0, 3, mathTop.y);
                                    attribute.SetValue(string.Empty, 0, 4, sizeX);
                                    attribute.SetValue(string.Empty, 0, 5, sizeY);
                                }*/

                        }
                    }
                    TempGroup.DetachObjects(TempGroup.Objects[0], true);
                    api.Doc.GiveLayer(0).Current = true;
                    api.Doc.GiveLayer(0).Update();
                }

                void SetText(string text, double x, double y, ksAllocationEnum alignEnum, double sizeText, double angle = 0, bool under = false)
                {
                    IDrawingTexts drawingTexts = api.Doc.GetDrawingContainer().DrawingTexts;
                    IDrawingText drawingText = drawingTexts.Add();
                    IText txt = (IText)drawingText;

                    drawingText.X = x;
                    drawingText.Y = y;
                    drawingText.Angle = angle;
                    drawingText.Allocation = ksAllocationEnum.ksAlCentre;

                    ITextLine textLine = txt.Add();
                    ITextItem textItem = textLine.Add();
                    textItem.Str = text;
                    ITextFont font = (ITextFont)textItem;
                    font.Italic = true;
                    font.Height = sizeText;
                    font.Underline = under;
                    textItem.Update();

                    drawingText.Update();
                }

                void SetPointToObj(object sendObj, ksMathPointParam mathBop, ksMathPointParam mathTop, double width)
                {
                    IDrawingObject pDrawObj = (IDrawingObject)sendObj;
                    if (pDrawObj != null) 
                    { 
                    double x = mathBop.x;
                    double y = mathBop.y;
                    double dx = mathTop.x;
                    double dy = mathTop.y;
                    double sizeX = Math.Round(Math.Abs(x - dx), 2);
                    double sizeY = Math.Round(Math.Abs(y - dy), 2);
                    double cx = x + sizeX / 2;
                    double cy = y + sizeY / 2;

                        if (pDrawObj != null)
                        {
                            // Получить тип объекта
                            long type = (int)pDrawObj.DrawingObjectType;

                            // В зависимости от типа вывести сообщение для данного типа объектов
                            switch (type)
                            {
                                // Линия Ок
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLineSeg:
                                    {
                                        ILineSegment obj = (ILineSegment)pDrawObj;
                                        if ((obj.Style == 1) || (obj.Style == 7))
                                        {
                                            bool invertX = Math.Abs(obj.X2 - cx) > Math.Abs(obj.X1 - cx);
                                            bool invertY = Math.Abs(obj.Y2 - cy) > Math.Abs(obj.Y1 - cy);

                                            SetPoint(obj.X1, obj.Y1, invertX, invertY);
                                            SetPoint(obj.X2, obj.Y2, !invertX, !invertY);

                                            LineSubcribe(obj.X1, obj.Y1, obj.X2, obj.Y2);
                                        }
                                        break;
                                    }
                                // Квадрат
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRectangle:
                                    {
                                        IRectangle obj = (IRectangle)pDrawObj;
                                        SetPoint(obj.X, obj.Y);
                                        SetPoint(obj.X, obj.Y + obj.Height);
                                        SetPoint(obj.X + obj.Width, obj.Y);
                                        SetPoint(obj.X + obj.Width, obj.Y + obj.Height);
                                        LineSubcribe(obj.X, obj.Y, obj.X, obj.Y + obj.Height);
                                        LineSubcribe(obj.X + obj.Width, obj.Y, obj.X + obj.Width, obj.Y + obj.Height);
                                        break;
                                    }
                                // Окружность
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrEllipse:
                                    {
                                        IEllipse obj = (IEllipse)pDrawObj;
                                        // if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += api.Mat.ksGetCurvePerimeter(objRef, 1);
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrCircle:
                                    {
                                        ICircle obj = (ICircle)pDrawObj;
                                        //if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += api.Mat.ksGetCurvePerimeter(objRef, 1);
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrEllipseArc:
                                    {
                                        IEllipseArc obj = (IEllipseArc)pDrawObj;
                                        //if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += api.Mat.ksGetCurvePerimeter(objRef, 1);
                                        break;
                                    }
                                //Nurbs
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrNurbs:
                                    {
                                        INurbs obj = (INurbs)pDrawObj;
                                        if ((obj.Style == 1) || (obj.Style == 7))
                                        {
                                            for (int k = 0; k < obj.PointsCount; k++)
                                            {
                                                double tx, ty, tw;
                                                obj.GetPoint(k, out tx, out ty, out tw);
                                                SetPoint(tx, ty);
                                            }
                                        }
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrNurbsByPoints:
                                    {
                                        INurbsByPoints obj = (INurbsByPoints)pDrawObj;
                                        if ((obj.Style == 1) || (obj.Style == 7))
                                        {
                                            ksDynamicArray arrayCurve = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(ldefin2d.POINT_ARR);
                                            ksDynamicArray arrs = (ksDynamicArray)api.Mat.ksPointsOnCurve(obj.Reference, (Int32)(api.Mat.ksGetCurvePerimeter(obj.Reference, 1) / 30));
                                            int count = arrs.ksGetArrayCount();
                                            for (int j = 0; j < count; j++)
                                            {
                                                ksMathPointParam ksPoints = (ksMathPointParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
                                                arrs.ksGetArrayItem(j, ksPoints);
                                                SetPoint(ksPoints.x, ksPoints.y);
                                            }
                                        }
                                        break;
                                    }

                                // Дуга 
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrArc:
                                    {
                                        IArc obj = (IArc)pDrawObj;
                                        if ((obj.Style == 1) || (obj.Style == 7))
                                        {
                                            SetPoint(obj.X1, obj.Y1);
                                            SetPoint(obj.X2, obj.Y2);
                                            SetPoint(obj.X3, obj.Y3);
                                        }

                                        break;
                                    }
                                //Безье
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBezier:
                                    {
                                        IBezier obj = (IBezier)pDrawObj;
                                        if ((obj.Style == 1) || (obj.Style == 7))
                                        {
                                            Array arrs = (Array)obj.Points[true];
                                        }
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrContour:
                                    {
                                        IContour obj = (IContour)pDrawObj;
                                        break;
                                    }
                            }
                        }


                    void SetPoint(double objX, double objY, bool invertX = false, bool invertY = false)
                    {
                            double textSize = width / 50;
                        Point point = new Point(objX, objY);
                            if (points.IndexOf(point) == -1)
                            {

                                //Создаем выноску
                                ViewsAndLayersManager ViewsMng = api.Doc.D7.ViewsAndLayersManager;
                                IViews views = ViewsMng.Views;
                                IView view = views.ActiveView;

                                IDrawingContainer drawingContainer = (IDrawingContainer)view;
                                ISymbols2DContainer symbols = (ISymbols2DContainer)view;
                                IBaseLeader baseLeader = symbols.Leaders.Add(DrawingObjectTypeEnum.ksDrLeader);
                                baseLeader.ArrowType = ksArrowEnum.ksLeaderPoint;

                                IBranchs branchs = (IBranchs)baseLeader;
                                branchs.AddBranchByPoint(0, objX, objY);
                                branchs.BranchX[0] = objX;
                                branchs.BranchY[0] = objY;
                                branchs.X0 = objX + (objX <= cx ? -width / 80 : width / 80);
                                branchs.Y0 = objY + (objY <= cy ? -width / 80 * 2 : width / 80);

                                ILeader leader = (ILeader)baseLeader;
                                leader.SignType = ksLeaderSignEnum.ksLSignNone;
                                IText textOnShelf = leader.TextOnShelf;
                                ITextLine BrachText = textOnShelf.Add();

                                leader.ShelfDirection = (objX <= cx ? ksShelfDirectionEnum.ksLSLeft : ksShelfDirectionEnum.ksLSRight);

                                if (invertX)
                                    leader.ShelfDirection = (leader.ShelfDirection == ksShelfDirectionEnum.ksLSLeft ? ksShelfDirectionEnum.ksLSRight : ksShelfDirectionEnum.ksLSLeft);
                                if (invertY)
                                    branchs.Y0 = objY + (invertY ? -width / 80 * 3 : width / 80 * 3);

                                //Проверяем X
                                if (NotNear(objX, objY, true, CoordDopusk))
                                {
                                    if ((Math.Round(objX, 0) != Math.Round(x, 0)) && (Math.Round(objX, 0) != Math.Round(dx, 0)))
                                    {
                                        ITextItem textItem = BrachText.Add();
                                        textItem.Str = (objX <= cx ? Math.Round(objX - x, 1) : Math.Round(dx - objX, 1)).ToString();
                                        textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                        ITextFont font = (ITextFont)textItem;
                                        font.Height = textSize;
                                        textItem.Update();

                                        ITextItem textItemUP = BrachText.Add();
                                        textItemUP.Str = "x";
                                        textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                        ITextFont fontUP = (ITextFont)textItem;
                                        font.Height = textSize;
                                        textItemUP.Update();

                                    }
                                }
                                else
                                {
                                    if (leader.ShelfDirection == ksShelfDirectionEnum.ksLSLeft) leader.ShelfDirection = ksShelfDirectionEnum.ksLSRight;
                                    else leader.ShelfDirection = ksShelfDirectionEnum.ksLSLeft; 
                                }
                                //Проверяем Y
                                if (NotNear(objY, objX, false, CoordDopusk))
                                {
                                    if ((Math.Round(objY, 0) != Math.Round(y, 0)) && (Math.Round(objY, 0) != Math.Round(dy, 0)))
                                    {
                                        ITextItem textItem = BrachText.Add();
                                        textItem.Str = (objY <= cy ? Math.Round(objY - y, 1) : Math.Round(dy - objY, 1)).ToString();
                                        textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                        ITextFont font = (ITextFont)textItem;
                                        font.Height = textSize;
                                        textItem.Update();

                                        ITextItem textItemUP = BrachText.Add();
                                        textItemUP.Str = "y";
                                        textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                        ITextFont fontUP = (ITextFont)textItem;
                                        font.Height = textSize;
                                        textItemUP.Update();
                                    }
                                }
                                else
                                {
                                    branchs.Y0 -= textSize;
                                }
                                points.Add(point);

                                if (BrachText.Str == string.Empty) 
                                    baseLeader.Delete();
                                else
                                    baseLeader.Update();
                            }
                    }

                    void LineSubcribe(double pointX1, double pointY1, double pointX2, double pointY2)
                    {

                        double distance = api.Mat.ksDistancePntPnt(pointX1, pointY1, pointX2, pointY2);

                        if (distance > width * 0.01)
                        {
                            double angle = api.Mat.ksAngle(pointX1, pointY1, pointX2, pointY2);
                            if ((angle >= 110) && (angle < 300)) angle -= 180;

                            IDrawingTexts drawingTexts = api.Doc.GetDrawingContainer().DrawingTexts;
                            IDrawingText drawingText = drawingTexts.Add();
                            IText txt = (IText)drawingText;

                            drawingText.X = pointX1 + (pointX2 - pointX1) / 2;
                            drawingText.Y = pointY1 + (pointY2 - pointY1) / 2;
                            drawingText.Allocation = ksAllocationEnum.ksAlCentre;
                            drawingText.Angle = angle;
                            ITextLine textLine = txt.Add();
                            textLine.Align = ksAlignEnum.ksAlignCenter;


                            ITextItem textItem = textLine.Add();
                            textItem.Str = Math.Round(distance, 0).ToString();
                            textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                            ITextFont font = (ITextFont)textItem;
                            font.Height = distance > width * 0.1 ? width / 30 : width / (30 * 2);
                            textItem.Update();
                            drawingText.Update();
                        }
                    }


                    bool NotNear(double Current, double Near, bool xflag, double distance)
                    {
                        for (int i = 0; i < points.Count; i++)

                            if (xflag)
                            {
                                if ((Math.Round(points[i].X, 0) == Math.Round(Current, 0)) && (points[i].Y + distance > Near) && (points[i].Y - distance < Near))
                                    return false;
                            }
                            else
                            {
                                if ((Math.Round(points[i].Y, 0) == Math.Round(Current, 0)) && (points[i].X + distance > Near) && (points[i].X - distance < Near))
                                    return false;
                            }
                        return true;
                    }
                }
                }
            };

            bool IndexOfTrue(int[] arr, int value)
            {
                for (int i = 0; i < arr.Length; i++)
                    if (arr[i] == value) return true;
                return false;
            }

        }


        //Проставить линейный размер
        public reference SetLineDim(double X1, double Y1, double X2, double Y2, double height, ksLineDimensionOrientationEnum orientationEnum = ksLineDimensionOrientationEnum.ksLinDParallel)
        {         
            if (api.Mat.ksEqualPoints(X1, Y1, X2, Y2) == 0)
            {
                api.Doc.VisibleLayer(77, false);
                api.Doc.VisibleLayer(88, true);

                ILayer layer = api.Doc.GiveLayer(88);

                layer.Color = 8355711;
                layer.Update();

                ViewsAndLayersManager ViewsMng = api.Doc.D7.ViewsAndLayersManager;
                IViews views = ViewsMng.Views;
                IView view = views.ActiveView;

                IDrawingContainer drawingContainer = (IDrawingContainer)view;
                ISymbols2DContainer symbols = (ISymbols2DContainer)view;
                ILineDimension lineDimension = symbols.LineDimensions.Add();
                

                IDimensionText dimensionText = (IDimensionText)lineDimension;
                IDimensionParams dimensionParams = (IDimensionParams)lineDimension;

                //Параметры размера
                lineDimension.X1 = X1;
                lineDimension.Y1 = Y1;
                lineDimension.X2 = X2;
                lineDimension.Y2 = Y2;
                //Вычисляем координату полки размера
                Point point = RotatePoint(new Point(X1, Y1), new Point(X2, Y2), height);
                lineDimension.X3 = point.X;
                lineDimension.Y3 = point.Y;

                lineDimension.Orientation = orientationEnum;


                //Параметры текста
                dimensionText.AutoNominalValue = true; //автоматический размер
                dimensionText.ToleranceOn = false; //Убираем квалитеты
                dimensionText.DeviationOn = false; //Погрешности
                                                   //Префикс

                //Параметры оформления
                dimensionParams.ArrowPos = ksDimensionArrowPosEnum.ksDimArrowInside;
                dimensionParams.ArrowType1 = ksArrowEnum.ksLeaderWithoutArrow;
                dimensionParams.ArrowType2 = ksArrowEnum.ksLeaderWithoutArrow;
                dimensionParams.TextBase = ksDimensionBaseEnum.ksDimBaseCenter;

                lineDimension.LayerNumber = layer.LayerNumber;
                lineDimension.Update();

                api.Doc.GiveLayer(0).Current = true;
                api.Doc.GiveLayer(0).Update();

                return lineDimension.Reference;


                Point RotatePoint(Point startPoint, Point endPoint, double hght, double angleInDegrees = 90)
                {
                    startPoint = startPoint + (endPoint - startPoint) / 2;
                    double lenth = api.Mat.ksDistancePntPnt(startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);
                    double pos = hght / lenth;
                    Point pointH = startPoint + (endPoint - startPoint) * pos;
                    double angleInRadians = angleInDegrees * (Math.PI / 180);
                    double cosTheta = Math.Cos(angleInRadians);
                    double sinTheta = Math.Sin(angleInRadians);
                    return new System.Windows.Point
                    {
                        X =
                            (double)
                            (cosTheta * (pointH.X - startPoint.X) -
                            sinTheta * (pointH.Y - startPoint.Y) + startPoint.X),
                        Y =
                            (double)
                            (sinTheta * (pointH.X - startPoint.X) +
                            cosTheta * (pointH.Y - startPoint.Y) + startPoint.Y)
                    };
                }
            }

            return 0;
        }

        /// <summary>
        /// Образмерить объект
        /// </summary>
        /// <returns>force — образмерить насильно, игнорируя настройки</returns>
        public void SizeMe(int ObjRef, bool SizeCheks, bool force = false)
        {
            if (SizeCheks || force == true)
            {
                if (KmpsAppl.KompasAPI != null)
                    GetSize(true);
                else
                    return;
            }
            //Расставить размеры в зависимости от типа объекта
            void GetSize(bool noSqare)
            {
                IDrawingObject pDrawObj = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(ObjRef, api.Doc.D5.reference);
                if (pDrawObj != null)
                {
                    // Получить тип объекта
                    long type = (int)pDrawObj.DrawingObjectType;

                    // Подсветить объект
                    api.Doc.D5.ksLightObj(ObjRef, 1/*включить*/ );

                    // В зависимости от типа вывести сообщение для данного типа объектов
                    switch (type)
                    {
                        // Линия
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLineSeg:
                            {
                                ILineSegment obj = (ILineSegment)pDrawObj;
                                if (obj.Style == 1 || obj.Style == 7)
                                {
                                    reference DimRef = SetLineDim(obj.X1, obj.Y1, obj.X2, obj.Y2, 20);
                                    SetConstrainttDim(DimRef, obj.Reference, 0, 1);
                                }
                                break;
                            }
                        // Квадрат
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRectangle:
                            {
                                IRectangle obj = (IRectangle)pDrawObj;
                                if (obj.Style == 1 || obj.Style == 7)
                                {
                                    reference DimRef = SetLineDim(obj.X + obj.Width, obj.Y + obj.Height, obj.X, obj.Y, 20, ksLineDimensionOrientationEnum.ksLinDHorizontal);
                                    SetConstrainttDim(DimRef, obj.Reference, 1, 0);
                                    DimRef = SetLineDim(obj.X, obj.Y, obj.X + obj.Width, obj.Y + obj.Height, 20, ksLineDimensionOrientationEnum.ksLinDVertical);
                                    SetConstrainttDim(DimRef, obj.Reference, 0, 1);
                                }
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrFragment:
                            {
                                ksFragment obj = (ksFragment)pDrawObj;
                                break;
                            }
                        // Окружность
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrEllipse:
                            {
                                IEllipse obj = (IEllipse)pDrawObj;
                                if (api.Doc.D5.ksMakeEncloseContours(0, obj.X1, obj.Y1) > 0)
                                {

                                }
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrCircle:
                            {
                                ICircle obj = (ICircle)pDrawObj;
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrEllipseArc:
                            {
                                IEllipseArc obj = (IEllipseArc)pDrawObj;
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrNurbs:
                            {
                                INurbs obj = (INurbs)pDrawObj;
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrNurbsByPoints:
                            {
                                INurbsByPoints obj = (INurbsByPoints)pDrawObj;
                                break;
                            }

                        // Дуга
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrArc:
                            {
                                IArc obj = (IArc)pDrawObj;
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBezier:
                            {
                                IBezier obj = (IBezier)pDrawObj;
                                break;
                            }
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrContour:
                            {
                                IContour obj = (IContour)pDrawObj;
                                break;
                            }
                    }
                    // Убрать подсветку
                    api.Doc.D5.ksLightObj(ObjRef, 0/*выключить*/ );
                    ILayer layer = api.Doc.GiveLayer(0);
                }

                //Привязка размера
                void SetConstrainttDim(reference dimRef, reference objRef, int Index1, int Index2)
                {
                    object DrawObj = (object)KmpsAppl.KompasAPI.TransferReference(objRef, api.Doc.D5.reference);
                    object Dim = (object)KmpsAppl.KompasAPI.TransferReference(dimRef, api.Doc.D5.reference);
                    IDrawingObject1 Dim1 = (IDrawingObject1)Dim;

                    //Накладываем на объект ограничение совпадение точек 1
                    IParametriticConstraint FixPoint1 = Dim1.NewConstraint();
                    if (FixPoint1 != null)
                    {
                        FixPoint1.Comment = "FixPoint1";
                        FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                        FixPoint1.Partner = DrawObj;
                        FixPoint1.Index = 0;
                        FixPoint1.PartnerIndex = Index1;
                        bool flag = FixPoint1.Create();
                    }

                    IParametriticConstraint FixPoint2 = Dim1.NewConstraint();
                    if (FixPoint2 != null)
                    {
                        FixPoint2.Comment = "FixPoint2";
                        FixPoint2.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                        FixPoint2.Partner = DrawObj;
                        FixPoint2.Index = 1;
                        FixPoint2.PartnerIndex = Index2;
                        bool flag = FixPoint2.Create();
                    }

                    //Накладываем на объект ограничение "Фиксированный размер"
                    IParametriticConstraint FixetDim = Dim1.NewConstraint();
                    if (FixetDim != null)
                    {
                        FixetDim.Comment = "FixConstraint";
                        FixetDim.ConstraintType = ksConstraintTypeEnum.ksCFixedDim;
                        if (!FixetDim.Create()) FixetDim.Delete();
                    }


                    //Накладываем на объект ограничение "Фиксированный размер"
                    IParametriticConstraint WithVariable = Dim1.NewConstraint();
                    if (WithVariable != null)
                    {
                        WithVariable.Comment = "WithVariable";
                        WithVariable.ConstraintType = ksConstraintTypeEnum.ksCDimWithVariable;
                        if (!WithVariable.Create()) WithVariable.Delete();

                    }

                }
            }
        }

        public void CreateRoom(double nSides, double nSideLength)
        {
            api.someFlag = false;
            ILineSegments lineSegments = api.Doc.GetDrawingContainer().LineSegments;
            Point LastPoint = new Point(0, 0);
            List<ILineSegment> Lines = new List<ILineSegment>();

            var step = 360.0 / (nSides);

            for (int i = 0; i < nSides; i++)
            {
                var deg = step + (180.0 * (nSides - 2)) / nSides - step * i + 90;
                var rad = deg * (Math.PI / 180);

                double nSinDeg = Math.Sin(rad);
                double nCosDeg = Math.Cos(rad);

                ILineSegment line = lineSegments.Add();

                line.X1 = LastPoint.X;
                line.Y1 = LastPoint.Y;
                line.X2 = LastPoint.X - nCosDeg * nSideLength;
                line.Y2 = LastPoint.Y - nSinDeg * nSideLength;
                line.Update();

                LastPoint.X = line.X2;
                LastPoint.Y = line.Y2;

                Lines.Add(line);
            }
            Lines.Add(Lines[0]);
            //Привязываем линии друг к другу
            for (int i = 0; i < Lines.Count - 1; i++)
                SetConstraintPoint(Lines[i].Reference, Lines[i + 1].Reference, 1, 0, i == 0 ? true : false);
            //образмериваем
           for (int i = 0; i < Lines.Count - 1; i++)
                SizeMe(Lines[i].Reference, true, true);


            api.someFlag = true;

            api.Doc.GiveLayer(0).Current = true;
            api.Doc.GiveLayer(0).Update();
            //Привязка размера
            void SetConstraintPoint(reference DimRef, reference ObjRef, int Index1, int Index2, bool first)
            {
                object line2 = (object)KmpsAppl.KompasAPI.TransferReference(ObjRef, api.Doc.D5.reference);
                object line1 = (object)KmpsAppl.KompasAPI.TransferReference(DimRef, api.Doc.D5.reference);
                IDrawingObject1 line1_1 = (IDrawingObject1)line1;

                //Накладываем на объект ограничение совпадение точек 1
                IParametriticConstraint FixPoint1 = line1_1.NewConstraint();
                if (FixPoint1 != null)
                {
                    FixPoint1.Comment = "FixPoint1";
                    FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                    FixPoint1.Partner = line2;
                    FixPoint1.Index = Index1;
                    FixPoint1.PartnerIndex = Index2;
                    FixPoint1.Create();
                }

                if (first) //Если первая прямая, то накладываем ограничения
                {
                    //Фиксируем точку А
                    IParametriticConstraint FixLine = line1_1.NewConstraint();
                    if (FixLine != null)
                    {
                        FixLine.Comment = "BaseLine";
                        FixLine.ConstraintType = ksConstraintTypeEnum.ksCFixedPoint;
                        FixLine.Index = 0;
                        FixLine.Create();
                    }
                    //Фиксируем точку А
                    IParametriticConstraint Vertical = line1_1.NewConstraint();
                    if (Vertical != null)
                    {
                        Vertical.Comment = "BaseLine";
                        Vertical.ConstraintType = ksConstraintTypeEnum.ksCVertical;
                        Vertical.Index = 0;
                        Vertical.Create();
                    }
                }
            }
        }

        public void SplitLine(double nSides, bool dellBaseSim)
        {
            api.someFlag = false;
            ISelectionManager selection = api.Doc.GetSelectContainer();

            if (selection.SelectedObjects != null)
            {
                double x = 0, y = 0;
                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                api.Doc.D5.ksCursor(info, ref x, ref y, 0);

                // Получить массив объектов
                try
                {
                    Array arrS = (Array)selection.SelectedObjects;
                    foreach (object obj in arrS) SplitLine(obj);
                }
                catch
                {
                    //если один объект
                    SplitLine(selection.SelectedObjects);
                }

                void SplitLine(object SpliObj)
                {
                    if (SpliObj != null)
                    {
                        IDrawingObject drawingObject = (IDrawingObject)SpliObj;

                        if (drawingObject != null)
                        {
                            if (drawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                            {
                                ILineSegment lineSegment = (ILineSegment)drawingObject;
                                if (lineSegment != null)
                                {
                                    IDrawingObject1 Constrait = (IDrawingObject1)SpliObj;

                                    List<ILineSegment> Lines = new List<ILineSegment>();//лист для линий.
                                    ILineSegment[] segmentsArr = new ILineSegment[2];
                                    foreach (IParametriticConstraint constraint in Constrait.Constraints)
                                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints)
                                        {
                                            try
                                            {
                                                //Берем партнеров
                                                Array array = (Array)constraint.Partner;
                                                //Переводив в объекты

                                                for (int i = 0; i < array.Length; i++)
                                                {
                                                    IDrawingObject drawingObject1 = (IDrawingObject)array.GetValue(i);
                                                    api.Doc.D5.ksLightObj(drawingObject1.Reference, 1/*включить*/ );
                                                    if (drawingObject1.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                                                    {
                                                        ILineSegment lineSegments1 = (ILineSegment)drawingObject1;
                                                        segmentsArr[constraint.Index] = lineSegments1;
                                                    }
                                                    api.Doc.D5.ksLightObj(drawingObject1.Reference, 0/*включить*/ );
                                                }


                                            }
                                            catch
                                            {
                                                //если только один объект то пока не нужно
                                            }
                                        }

                                    #region Создание линий"
                                    ILineSegments segments = api.Doc.GetDrawingContainer().LineSegments;

                                    Lines.Add(segmentsArr[0]); //добавляем в список предыдущий

                                    //Начальная конечная точка
                                    System.Windows.Point start = new System.Windows.Point(lineSegment.X1, lineSegment.Y1);
                                    System.Windows.Point end = new System.Windows.Point(lineSegment.X2, lineSegment.Y2);
                                    //Удаляем линию
                                    lineSegment.Delete();

                                    ILineSegments lineSegments = api.Doc.GetDrawingContainer().LineSegments;
                                    Point LastPoint = new Point(0, 0);

                                    for (int i = 0; i < nSides; i++)
                                    {
                                        ILineSegment line = lineSegments.Add();

                                        line.X1 = start.X + ((end.X - start.X) / nSides * i);
                                        line.Y1 = start.Y + ((end.Y - start.Y) / nSides * i);
                                        line.X2 = start.X + ((end.X - start.X) / nSides * (i + 1));
                                        line.Y2 = start.Y + ((end.Y - start.Y) / nSides * (i + 1));
                                        line.Update(); //Апдейтим

                                        Lines.Add(line); //Добавляем в писок линию
                                    }


                                    Lines.Add(segmentsArr[1]); //добавляем в список следующую

                                    //Связываем линии
                                    for (int i = 0; i < Lines.Count - 1; i++)
                                        SetConstraintPoint(Lines[i].Reference, Lines[i + 1].Reference, 1, 0);

                                    //образмериваем информационными
                                    for (int i = 1; i < Lines.Count - 1; i++)
                                        SetConstrainttDim(SetLineDim(Lines[i].X1, Lines[i].Y1, Lines[i].X2, Lines[i].Y2, 20), Lines[i].Reference, 0, 1, ksDimensionTextBracketsEnum.ksDimBrackets, true, false);

                                    #endregion
                                    //образмериваем по точке
                                        for (int i = 1; i < Lines.Count - 1; i++)
                                            SetConstrainttDim(SetLineDim(x, y, Lines[i].X2, Lines[i].Y2, 0), Lines[i].Reference, 0, 1, ksDimensionTextBracketsEnum.ksDimBracketsOff, false, true);
                                }
                            }


                            ///////////////процедуры///////////////////

                            //Привязка размера
                            void SetConstrainttDim(reference DimRef, reference ObjRef, int Index1, int Index2, ksDimensionTextBracketsEnum bracketsEnum, bool infoDim, bool toPoint = false)
                            {
                                object DrawObj = (object)KmpsAppl.KompasAPI.TransferReference(ObjRef, api.Doc.D5.reference);
                                object Dim = (object)KmpsAppl.KompasAPI.TransferReference(DimRef, api.Doc.D5.reference);
                                if (Dim != null)
                                {
                                    IDrawingObject objDim = (IDrawingObject)Dim;
                                    ILineDimension lineDimension = (ILineDimension)objDim;
                                    IDimensionText dimensionText = (IDimensionText)lineDimension;
                                    dimensionText.Brackets = bracketsEnum;
                                    lineDimension.Update();

                                    IDrawingObject1 Dim1 = (IDrawingObject1)Dim;

                                    IParametriticConstraint FixPoint1 = Dim1.NewConstraint();
                                    //Фиксируем точку на координатах или объекте
                                    if (toPoint)
                                    {
                                        if (FixPoint1 != null)
                                        {
                                            FixPoint1.Comment = "FixPoint1";
                                            FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCFixedPoint;
                                            FixPoint1.Partner = DrawObj;
                                            FixPoint1.Index = 0;
                                        }
                                    }
                                    else
                                    {
                                        if (FixPoint1 != null)
                                        {
                                            FixPoint1.Comment = "FixPoint1";
                                            FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                                            FixPoint1.Partner = DrawObj;
                                            FixPoint1.Index = 0;
                                            FixPoint1.PartnerIndex = Index1;

                                        }
                                    }
                                    bool flag = FixPoint1.Create();


                                    //Фиксируем на объекте
                                    IParametriticConstraint FixPoint2 = Dim1.NewConstraint();
                                    if (FixPoint2 != null)
                                    {
                                        FixPoint2.Comment = "FixPoint2";
                                        FixPoint2.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                                        FixPoint2.Partner = DrawObj;
                                        FixPoint2.Index = 1;
                                        FixPoint2.PartnerIndex = Index2;
                                        FixPoint2.Create();
                                    }

                                    if (!infoDim)
                                        SetParamToObj(Dim, false, bracketsEnum);
                                }
                            }

                            //Привязка точки объектов
                            void SetConstraintPoint(reference Obj1Ref, reference Obj2Ref, int Index1, int Index2)
                            {
                                api.Doc.D5.ksLightObj(Obj1Ref, 1/*включить*/ );
                                object line2 = (object)KmpsAppl.KompasAPI.TransferReference(Obj2Ref, api.Doc.D5.reference);
                                object line1 = (object)KmpsAppl.KompasAPI.TransferReference(Obj1Ref, api.Doc.D5.reference);
                                IDrawingObject1 line1_1 = (IDrawingObject1)line1;

                                //Накладываем на объект ограничение совпадение точек 1
                                IParametriticConstraint FixPoint1 = line1_1.NewConstraint();
                                if (FixPoint1 != null)
                                {
                                    FixPoint1.Comment = "FixPoint1";
                                    FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                                    FixPoint1.Partner = line2;
                                    FixPoint1.Index = Index1;
                                    FixPoint1.PartnerIndex = Index2;
                                    FixPoint1.Create();
                                }
                                api.Doc.D5.ksLightObj(Obj1Ref, 0/*включить*/ );
                            }
                        }
                    }
                }

            }

            api.someFlag = true;
        }

        public void SetVariableToDim(bool dell, ksDimensionTextBracketsEnum brackets, double VariableValue = 0)
        {
            ISelectionManager selectionManager = api.Doc.GetSelectContainer();
            if (selectionManager != null)
            {
                try
                {
                    //Перебираем объекты. Определяем что за объект внутри процедуры
                    Array array = (Array)selectionManager.SelectedObjects;
                    foreach (object obj in array) 
                        if (obj != null)
                            SetParamToObj(obj, dell, brackets, VariableValue);
                }
                catch
                {
                    SetParamToObj((object)selectionManager.SelectedObjects, dell, brackets, VariableValue);
                }
            }
           
        }
        private void SetParamToObj(object objDim, bool dell, ksDimensionTextBracketsEnum bracketsEnum, double VariableValue = 0)
        {
            IDrawingObject SetDrawingObject = (IDrawingObject)objDim;
            if (SetDrawingObject != null)
            //Определяем что это линия
            if (SetDrawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrLDimension)
            {
                ILineDimension lineDimension = (ILineDimension)SetDrawingObject;
                IDimensionText dimensionText = (IDimensionText)lineDimension;
                IDrawingObject1 draw1 = (IDrawingObject1)objDim;

                if (dell)
                {
                    if (draw1.Constraints != null)
                        foreach (IParametriticConstraint constraint in draw1.Constraints)
                            if ((constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable) ||
                                (constraint.ConstraintType == ksConstraintTypeEnum.ksCFixedDim)) constraint.Delete();
                }
                else
                {
                    if (draw1.Constraints != null)
                        foreach (IParametriticConstraint constraint in draw1.Constraints)
                            if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                                if (VariableValue == 0)
                                {
                                    if (VariableValue != 0) SetValToVariableDim();
                                    dimensionText.Brackets = bracketsEnum;
                                    lineDimension.Update();
                                    return;
                                }


                    //Накладываем на объект ограничение "Фиксированный размер"
                    IParametriticConstraint FixetDim = draw1.NewConstraint();
                    if (FixetDim != null)
                    {
                        FixetDim.Comment = "FixConstraint";
                        FixetDim.ConstraintType = ksConstraintTypeEnum.ksCFixedDim;
                        if (!FixetDim.Create())
                        { 
                            bracketsEnum = ksDimensionTextBracketsEnum.ksDimBrackets;
                            FixetDim.Delete();
                        }
                        
                    }

                    //Накладываем на объект ограничение "Размер с переменной"
                    IParametriticConstraint WithVariable = draw1.NewConstraint();
                    if (WithVariable != null)
                    {
                        WithVariable.Comment = "WithVariable";
                        WithVariable.ConstraintType = ksConstraintTypeEnum.ksCDimWithVariable;
                        if (!WithVariable.Create())
                        {
                            bracketsEnum = ksDimensionTextBracketsEnum.ksDimBrackets;
                            WithVariable.Delete();
                        }
                    }
                    if (VariableValue != 0) SetValToVariableDim();

                }
                dimensionText.Brackets = bracketsEnum;
                lineDimension.Update();


                //
                //Процедуры
                //
                void SetValToVariableDim()
                {
    

                    foreach (IParametriticConstraint constraint in draw1.Constraints)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                            api.Doc.Variable.Update(constraint.Variable, VariableValue, string.Empty);

                }
            }
        }

        public void LineToPoint()
        {
            api.someFlag = false;
            ISelectionManager selection = api.Doc.GetSelectContainer();

            if (selection.SelectedObjects != null)
            {
                List<Point> points = new List<Point>();
                // Получить массив объектов
                try
                {
                    Array arrS = (Array)selection.SelectedObjects;
                    foreach (object obj in arrS) GivePoint(obj);
                }
                catch
                {
                    //если один объект
                    GivePoint(selection.SelectedObjects);
                }

                #region Создание точек"
                List<IPoint> pointsList = new List<IPoint>();
                IPoints ipoints = api.Doc.GetDrawingContainer().Points;
                for (int i = 0; i < points.Count; i++)
                {
                    IPoint pnt = ipoints.Add();
                    pnt.X = points[i].X;
                    pnt.Y = points[i].Y;
                    pnt.Style = (int)ksAnnotationSymbolEnum.ksXPoint;
                    pnt.Update();
                    pointsList.Add(pnt);
                }

                //образмериваем информационными
                for (int i = 0; i < pointsList.Count - 1; i++)
                    SetConstrainttDim(SetLineDim(pointsList[i].X, pointsList[i].Y, pointsList[i + 1].X, pointsList[i + 1].Y, 20), pointsList[i].Reference, pointsList[i + 1].Reference, 0, 1, ksDimensionTextBracketsEnum.ksDimBrackets, true, false);

                 #endregion


                ///////////////процедуры///////////////////

                //Привязка размера
                void SetConstrainttDim(reference DimRef, reference Pnt1, reference Pnt2, int Index1, int Index2, ksDimensionTextBracketsEnum bracketsEnum, bool infoDim, bool toPoint = false)
                {
                    object pnt1 = (object)KmpsAppl.KompasAPI.TransferReference(Pnt1, api.Doc.D5.reference);
                    object pnt2 = (object)KmpsAppl.KompasAPI.TransferReference(Pnt2, api.Doc.D5.reference);
                    object Dim = (object)KmpsAppl.KompasAPI.TransferReference(DimRef, api.Doc.D5.reference);
                    if (Dim != null)
                    {
                        IDrawingObject objDim = (IDrawingObject)Dim;
                        ILineDimension lineDimension = (ILineDimension)objDim;
                        IDimensionText dimensionText = (IDimensionText)lineDimension;
                        dimensionText.Brackets = bracketsEnum;
                        lineDimension.Update();

                        IDrawingObject1 Dim1 = (IDrawingObject1)Dim;

                        IParametriticConstraint FixPoint1 = Dim1.NewConstraint();
                        //Фиксируем точку на координатах или объекте
                        if (FixPoint1 != null)
                        {
                            FixPoint1.Comment = "FixPoint1";
                            FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                            FixPoint1.Partner = pnt1;
                            FixPoint1.Index = 0;
                            FixPoint1.PartnerIndex = 0;
                            FixPoint1.Create();
                        }
                        IParametriticConstraint FixPoint2 = Dim1.NewConstraint();
                        if (FixPoint2 != null)
                        {
                            FixPoint2.Comment = "FixPoint2";
                            FixPoint2.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                            FixPoint2.Partner = pnt2;
                            FixPoint2.Index = 1;
                            FixPoint2.PartnerIndex = 0;
                            FixPoint2.Create();
                        }


                        if (!infoDim)
                            SetParamToObj(Dim, false, bracketsEnum);
                    }
                }

                void GivePoint(object SpliObj)
                {
                    if (SpliObj != null)
                    {
                        IDrawingObject drawingObject = (IDrawingObject)SpliObj;

                        if (drawingObject != null)
                        {
                            if (drawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                            {
                                ILineSegment lineSegment = (ILineSegment)drawingObject;
                                if (lineSegment != null)
                                {
                                    IDrawingObject1 Constrait = (IDrawingObject1)SpliObj;

                                    //Удаляем размеры
                                    foreach (IParametriticConstraint constraint in Constrait.Constraints)
                                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints)
                                        {
                                            try
                                            {
                                                //Берем партнеров
                                                Array array = (Array)constraint.Partner;
                                                //Переводив в объекты

                                                for (int i = 0; i < array.Length; i++)
                                                {
                                                    IDrawingObject drawingObject1 = (IDrawingObject)array.GetValue(i);
                                                    api.Doc.D5.ksLightObj(drawingObject1.Reference, 1/*включить*/ );
                                                    if (drawingObject1.DrawingObjectType == DrawingObjectTypeEnum.ksDrLDimension)
                                                    {
                                                        drawingObject1.Delete();
                                                    }
                                                    api.Doc.D5.ksLightObj(drawingObject1.Reference, 0/*включить*/ );
                                                }


                                            }
                                            catch
                                            {
                                                //если только один объект то пока не нужно
                                            }
                                        }
                                    //Добавляем точки
                                    if (points.IndexOf(new Point((double)lineSegment.X1, (double)lineSegment.Y1)) == -1)
                                        points.Add(new Point((double)lineSegment.X1, (double)lineSegment.Y1));
                                    if (points.IndexOf(new Point((double)lineSegment.X2, (double)lineSegment.Y2)) == -1)
                                        points.Add(new Point((double)lineSegment.X2, (double)lineSegment.Y2));

                                    //Удаляем линию
                                    lineSegment.Delete();


                                   }
                            }

                        }
                    }
                }

            }

            api.someFlag = true;
        }

        public void InvertPointCoord(bool xinvert)
        {
            api.someFlag = false;
            ISelectionManager selection = api.Doc.GetSelectContainer();

            if (selection.SelectedObjects != null)
            {
                try
                {
                    Array arrS = (Array)selection.SelectedObjects;

                    foreach (object obj in arrS)
                        invert(obj);

                }
                catch
                {
                    invert(selection.SelectedObjects);
                }

            }

            api.someFlag = true;

            void invert (object obj)
            {
                try 
                {
                    IBaseLeader pObj = (IBaseLeader)obj;
                    SetPoint(pObj, xinvert);
                }
                catch { return; }
                
                void SetPoint(IBaseLeader baseLeader, bool invertX)
                {
                    IAttribute attribute = (IAttribute)baseLeader.Parent;
                    if (attribute != null)
                    {
                        //Координаты области
                        double x = (double)attribute.Value[0, 0];
                        double y = (double)attribute.Value[0, 1];
                        double dx = (double)attribute.Value[0, 2];
                        double dy = (double)attribute.Value[0, 3];
                        //Ее размер
                        double sizeX = (double)attribute.Value[0, 4];
                        double sizeY = (double)attribute.Value[0, 5];

                        IBranchs branchs = (IBranchs)baseLeader;

                        double objX = branchs.BranchX[0];
                        double objY = branchs.BranchY[0];


                        ILeader leader = (ILeader)baseLeader;

                        IText textOnShelf = leader.TextOnShelf;
                        ITextLine BrachText = textOnShelf.TextLine[0];

                        //Проверяем X
                        if (FindOrd(BrachText, "x"))
                        {
                            if ((Math.Round(objX, 0) != Math.Round(x, 0)) && (Math.Round(objX, 0) != Math.Round(dx, 0)))
                            {
                                ITextItem textItem = BrachText.Add();
                                textItem.Str = "*" + (!invertX ? Math.Round(objX - x, 1) : Math.Round(dx - objX, 1)).ToString();
                                textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                textItem.Update();

                                ITextItem textItemUP = BrachText.Add();
                                textItemUP.Str = "x";
                                textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                textItemUP.Update();

                            }
                        }
                        else
                        {
                            if (leader.ShelfDirection == ksShelfDirectionEnum.ksLSLeft) leader.ShelfDirection = ksShelfDirectionEnum.ksLSRight;
                            else leader.ShelfDirection = ksShelfDirectionEnum.ksLSLeft;
                        }
                        //Проверяем Y
                        if (FindOrd(BrachText, "y"))
                        {
                            if ((Math.Round(objY, 0) != Math.Round(y, 0)) && (Math.Round(objY, 0) != Math.Round(dy, 0)))
                            {
                                ITextItem textItem = BrachText.Add();
                                textItem.Str = "*" + (invertX ? Math.Round(objY - y, 1) : Math.Round(dy - objY, 1)).ToString();
                                textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                ITextFont font = (ITextFont)textItem;

                                textItem.Update();

                                ITextItem textItemUP = BrachText.Add();
                                textItemUP.Str = "y";
                                textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                ITextFont fontUP = (ITextFont)textItem;

                                textItemUP.Update();
                            }
                        }

                        if (BrachText.Str == string.Empty)
                            baseLeader.Delete();
                        else
                            baseLeader.Update();
                    }

                    bool FindOrd(ITextLine textLine, string findStr)
                    {
                        foreach (ITextItem textItem in textLine.TextItems)
                            if (textItem.Str == findStr) return true;
                        return false;

                    }
                }
            }
        }

    }

}
