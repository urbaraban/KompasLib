using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using static KompasLib.KompasTool.HlpClasses;
using reference = System.Int32;

namespace KompasLib.Tools
{
    public class SizeTool
    {
        private readonly KmpsDoc doc;

        private List<ComboData> forDimList = new List<ComboData>();

        public static object SelectObj;
        public static bool SelectPointFlag = false;
        public bool WorkFlag = false;
        public static ksPhantom phtm;
        public static IDrawingGroup PhGroup;

        private static ILineDimension firstDim = null;
        private static readonly List<ILineSegment> lastLine = new List<ILineSegment>();

        private static readonly int[] YesType = { 1, 2, 3, 8, 26, 28, 31, 32, 33, 34, 35, 36, 80 };
        private static readonly int[] DimType = { 9, 10, 13, 14, 15, 43 };

        public SizeTool(KmpsDoc Doc)
        {
            this.doc = Doc;
            phtm = KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_Phantom);
        }
        public event EventHandler<List<List<ComboData>>> ChangeListDimention;

        /// <summary>
        /// Возвращает список размеров
        /// </summary>
        /// <returns></returns>
        public async Task<List<List<ComboData>>> GetLineDimentionListAsync()
        {
            List<List<ComboData>> comboDatas = new List<List<ComboData>>
            {
                new List<ComboData>(), //лист просто размеров
                new List<ComboData>(), //лист свободных размеров
                new List<ComboData>() //лист фиксированных размеров
            };


            if (KmpsAppl.KompasAPI != null)
            {
                ILineDimensions lineDimension = (ILineDimensions)this.doc.GetSymbols2DContainer().LineDimensions;

                var result = await Task<List<List<ComboData>>>.Factory.StartNew(() =>
            {
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
                                        comboDatas[0].Add(new ComboData("Размер " + i, dimension.Reference, comboDatas[0].Count));
                    }
                    else
                    {
                        if (dimensionText.Brackets == ksDimensionTextBracketsEnum.ksDimBrackets)
                            comboDatas[1].Add(new ComboData("Размер " + i, dimension.Reference, comboDatas[1].Count));
                        else
                            comboDatas[2].Add(new ComboData("Размер " + i, dimension.Reference, comboDatas[2].Count));
                    }

                }

                ChangeListDimention(this, comboDatas);

                return comboDatas;
            });
            }

            return null;
        }



        /// <summary>
        /// Описывает фигуру координатами
        /// </summary>
        /// <returns></returns>
        public async void Coordinate(double width, double CoordDopusk, double sizeText)
        {
            this.doc.GetChooseContainer().UnchooseAll();

            ILayer layer = this.doc.GetLayer(88);
            List<Point> points = new List<Point>();
            if (this.doc != null)
            {
                ISelectionManager selection = this.doc.GetSelectContainer();
                if (selection.SelectedObjects != null)
                {
                    IDrawingGroup TempGroup = (IDrawingGroup)KmpsAppl.KompasAPI.TransferReference(this.doc.D5.ksNewGroup(0), this.doc.D5.reference);
                    if (selection.SelectedObjects is Array array)
                    {
                        int i = 0;
                        KmpsAppl.ProgressBar.Start(i, array.Length, "Скрываем размеры", true);

                        bool tempflag = MessageBox.Show("Скрыть размеры?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
                        foreach (object obj in array)
                        {
                            await Task.Run(() =>
                            {
                                IDrawingObject pObj = (IDrawingObject)obj;
                                if (IndexOfTrue(YesType, (int)pObj.DrawingObjectType))
                                    TempGroup.AddObjects(obj);
                                else if (IndexOfTrue(DimType, (int)pObj.DrawingObjectType))
                                { pObj.LayerNumber = layer.LayerNumber; pObj.Update(); }
                            });
                        }
                        KmpsAppl.ProgressBar.Stop("Закончили", true);
                    }
                    else
                    {
                        TempGroup.AddObjects(selection.SelectedObjects);
                    }
                    this.doc.VisibleLayer(77, true);
                    this.doc.VisibleLayer(88, false);

                    TempGroup.Close();

                    ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                    ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                    if (this.doc.D5.ksGetObjGabaritRect(TempGroup.Reference, spcGabarit) == 1)
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
                        if (width == 0)
                            width = 320;

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

                        this.doc.D5.ksEndObj();

                        ViewsAndLayersManager ViewsMng = this.doc.D7.ViewsAndLayersManager;
                        IViews views = ViewsMng.Views;
                        IView view = views.ActiveView;

                        layer = this.doc.GetLayer(77);

                        layer.Color = 8355711;
                        layer.Update();


                        if (this.doc.D5.ksMacro(0) == 1)
                        {
                            //Вертикальная
                            this.doc.D5.ksLineSeg(x, y - kooff, x, dy + kooff, 3);
                            this.doc.D5.ksLineSeg(dx, y - kooff, dx, dy + kooff, 3);
                            //Горизонтальные
                            this.doc.D5.ksLineSeg(x - kooff, y, dx + kooff, y, 3);
                            this.doc.D5.ksLineSeg(x - kooff, dy, dx + kooff, dy, 3);
                            //Центр
                            this.doc.D5.ksLineSeg(x + (dx - x) / 2, y - kooff, x + (dx - x) / 2, dy + kooff, 3);
                            this.doc.D5.ksLineSeg(x - kooff, y + (dy - y) / 2, dx + kooff, y + (dy - y) / 2, 3);
                            //Ширина
                            SetText(Math.Round(mathTop.x - mathBop.x, 1, MidpointRounding.AwayFromZero).ToString(), mathBop.x + (mathTop.x - mathBop.x) / 2, mathBop.y - kooff * 1.5, ksAllocationEnum.ksAlCentre, 0, true);
                            //Высота
                            SetText(Math.Round(mathTop.y - mathBop.y, 1, MidpointRounding.AwayFromZero).ToString(), mathBop.x - kooff, mathBop.y + (mathTop.y - mathBop.y) / 2, ksAllocationEnum.ksAlCentre, 90, true);

                            //Подпись Y
                            SetText((!xflag ? "<<" : string.Empty) + "Y", x, y - kooff, ksAllocationEnum.ksAlLeft, -90);
                            //Подпись Х
                            SetText("X" + (xflag ? ">>" : string.Empty), x - kooff, y, ksAllocationEnum.ksAlRight);

                            ///
                            /// Идем по объектам и проставляем подписи
                            ///
                            if (TempGroup.Objects[0] is Array arrS) 
                            { 
                                KmpsAppl.ProgressBar.Start(0, arrS.Length, "Точка:", true);
                                for (int i = 0; i < arrS.Length; i++)
                                {
                                    KmpsAppl.ProgressBar.SetProgress(i, "Точка:", true);
                                    SetPointToObj(arrS.GetValue(i), mathBop, mathTop);
                                }
                                KmpsAppl.ProgressBar.Stop("Закончили", true);
                            }
                            else
                            {
                                SetPointToObj(TempGroup.Objects[0], mathBop, mathTop);
                            }

                            //конец макрообъекта
                            IMacroObject pMacroObj = (IMacroObject)KmpsAppl.KompasAPI.TransferReference(this.doc.D5.ksEndObj(), this.doc.D5.reference);
                            pMacroObj.Name = "CoordMacro:" + pMacroObj.Reference;
                            pMacroObj.LayerNumber = 77;
                            pMacroObj.Update();

                            TempGroup.DetachObjects(TempGroup.Objects[0], true);

                            IDrawingContainer drawingContainer = (IDrawingContainer)pMacroObj;

                            foreach (IDrawingObject drawingObject in drawingContainer.Objects[0])
                                await Task.Run(() =>
                                {
                                    IAttribute attribute = (IAttribute)KmpsAppl.KompasAPI.TransferReference(this.doc.Attribute.NewAttr(drawingObject.Reference), this.doc.D5.reference);
                                    if (attribute != null)
                                    {
                                        attribute.SetValue(string.Empty, 0, 0, mathBop.x);
                                        attribute.SetValue(string.Empty, 0, 1, mathBop.y);
                                        attribute.SetValue(string.Empty, 0, 2, mathTop.x);
                                        attribute.SetValue(string.Empty, 0, 3, mathTop.y);
                                        attribute.SetValue(string.Empty, 0, 4, sizeX);
                                        attribute.SetValue(string.Empty, 0, 5, sizeY);
                                    }
                                });
                            KmpsAppl.ProgressBar.Stop("Закончили", true);


                        }
                    }

                    this.doc.GetLayer(0).Current = true;
                    this.doc.GetLayer(0).Update();
                }

                void SetText(string text, double x, double y, ksAllocationEnum alignEnum, double angle = 0, bool under = false)
                {
                    IDrawingTexts drawingTexts = this.doc.GetDrawingContainer().DrawingTexts;
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

                void SetPointToObj(object sendObj, ksMathPointParam mathBop, ksMathPointParam mathTop)
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
                                        // if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += KmpsAppl.Mat.ksGetCurvePerimeter(objRef, 1);
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrCircle:
                                    {
                                        ICircle obj = (ICircle)pDrawObj;
                                        //if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += KmpsAppl.Mat.ksGetCurvePerimeter(objRef, 1);
                                        break;
                                    }
                                case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrEllipseArc:
                                    {
                                        IEllipseArc obj = (IEllipseArc)pDrawObj;
                                        //if ((obj.Style == 1) || (obj.Style == 7)) arcPeri += KmpsAppl.Mat.ksGetCurvePerimeter(objRef, 1);
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
                                                obj.GetPoint(k, out double tx, out double ty, out double tw);
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
                                            ksDynamicArray arrs = (ksDynamicArray)KmpsAppl.Mat.ksPointsOnCurve(obj.Reference, (Int32)(KmpsAppl.Mat.ksGetCurvePerimeter(obj.Reference, 1) / 30));
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


                        async void SetPoint(double objX, double objY, bool invertX = false, bool invertY = false)
                        {
                            double textSize = width / 50;
                            Point point = new Point(objX, objY);
                            if (points.IndexOf(point) == -1)
                            {
                                await Task.Run(() =>
                                {
                                    //Создаем выноску
                                    ViewsAndLayersManager ViewsMng = this.doc.D7.ViewsAndLayersManager;
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
                                            textItem.Str = (objX <= cx ? Math.Round(x - objX, 1) : Math.Round(dx - objX, 1)).ToString();
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
                                            textItem.Str = (objY <= cy ? Math.Round(y - objY, 1) : Math.Round(dy - objY, 1)).ToString();
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
                                });
                            }
                        }

                        void LineSubcribe(double pointX1, double pointY1, double pointX2, double pointY2)
                        {

                            double distance = KmpsAppl.Mat.ksDistancePntPnt(pointX1, pointY1, pointX2, pointY2);

                            if (distance > width * 0.01)
                            {
                                double angle = KmpsAppl.Mat.ksAngle(pointX1, pointY1, pointX2, pointY2);
                                if ((angle >= 110) && (angle < 300)) angle -= 180;

                                IDrawingTexts drawingTexts = this.doc.GetDrawingContainer().DrawingTexts;
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



        }

        //велосипед
        private static bool IndexOfTrue(int[] arr, int value)
        {
            for (int i = 0; i < arr.Length; i++)
                if (arr[i] == value) return true;
            return false;
        }

        /// <summary>
        /// Создает угловой размер между двумя линиями
        /// </summary>
        /// <returns>Возвращает класс IAngleDimention</returns>
        public IAngleDimension SetAngleDim(ILineSegment lineSegment1, ILineSegment lineSegment2)
        {
            this.doc.VisibleLayer(77, false);
            this.doc.VisibleLayer(88, true);
            _ = this.doc.GetLayer(88);

            ViewsAndLayersManager ViewsMng = this.doc.D7.ViewsAndLayersManager;
            IViews views = ViewsMng.Views;
            IView view = views.ActiveView;

            ISymbols2DContainer symbols = (ISymbols2DContainer)view;
            IAngleDimension angleDimension = symbols.AngleDimensions.Add(DrawingObjectTypeEnum.ksDrADimension);
            IDimensionParams dimensionParams = (IDimensionParams)angleDimension;

            angleDimension.Angle1 = lineSegment1.Angle;
            angleDimension.Angle2 = lineSegment2.Angle;
            angleDimension.BaseObject1 = lineSegment1;
            angleDimension.BaseObject2 = lineSegment2;
            angleDimension.DimensionType = ksAngleDimTypeEnum.ksADMinAngle;
            angleDimension.Direction = false;
            angleDimension.Radius = 30;

            angleDimension.LayerNumber = 88;

            dimensionParams.RemoteLine1 = false;
            dimensionParams.RemoteLine2 = false;

            angleDimension.Update();
            return angleDimension;
        }

        //Проставить линейный размер
        public ILineDimension SetLineDim(double X1, double Y1, double X2, double Y2, double height, ksLineDimensionOrientationEnum orientationEnum = ksLineDimensionOrientationEnum.ksLinDParallel)
        {

            if (KmpsAppl.Mat.ksEqualPoints(X1, Y1, X2, Y2) == 0)
            {
                this.doc.VisibleLayer(77, false);
                this.doc.VisibleLayer(88, true);

                ILayer layer = this.doc.GetLayer(88);
                layer.Color = 8355711;
                layer.Update();

                ISymbols2DContainer symbols = this.doc.GetSymbols2DContainer();
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
                dimensionText.TextAlign = ksDimensionTextAlignEnum.ksDimACentre;

                //Параметры оформления
                dimensionParams.ArrowPos = ksDimensionArrowPosEnum.ksDimArrowInside;
                dimensionParams.ArrowType1 = ksArrowEnum.ksLeaderWithoutArrow;
                dimensionParams.ArrowType2 = ksArrowEnum.ksLeaderWithoutArrow;
                dimensionParams.TextBase = ksDimensionBaseEnum.ksDimBaseCenter;

                lineDimension.LayerNumber = 88;
                lineDimension.Update();

                ITextLine textLine = dimensionText.NominalText;
                foreach (ITextItem textItem in textLine.TextItems)
                {
                    ITextFont textFont = (ITextFont)textItem;
                    textFont.Height = height;
                    textFont.Underline = true;
                }

                lineDimension.Update();

                this.doc.GetLayer(0);


                return lineDimension;

                Point RotatePoint(Point startPoint, Point endPoint, double hght, double angleInDegrees = 90)
                {
                    startPoint += (endPoint - startPoint) / 2;
                    double lenth = KmpsAppl.Mat.ksDistancePntPnt(startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);
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

            return null;
        }

        /// <summary>
        /// Образмерить объект
        /// </summary>
        /// <returns>force — образмерить насильно, игнорируя настройки</returns>
        public void SizeMe(int ObjRef, bool SizeCheks, bool variable, bool force = false)
        {
            if (SizeCheks || force == true)
            {
                if (KmpsAppl.KompasAPI != null)
                    GetSize();
                else
                    return;
            }
            //Расставить размеры в зависимости от типа объекта
            void GetSize()
            {
                IDrawingObject pDrawObj = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(ObjRef, this.doc.D5.reference);
                if (pDrawObj != null)
                {
                    // Получить тип объекта
                    long type = (int)pDrawObj.DrawingObjectType;
                    // В зависимости от типа вывести сообщение для данного типа объектов
                    switch (type)
                    {
                        // Линия
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLineSeg:
                            {
                                ILineSegment obj = (ILineSegment)pDrawObj;
                                if (obj.Style == 1 || obj.Style == 7)
                                {
                                    object Dim = SetLineDim(obj.X1, obj.Y1, obj.X2, obj.Y2, 20);
                                    SetConstrainttDim(Dim, obj, 0, 1, variable);
                                }
                                break;
                            }
                        // Квадрат
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRectangle:
                            {
                                IRectangle obj = (IRectangle)pDrawObj;
                                if (obj.Style == 1 || obj.Style == 7)
                                {
                                    object Dim = SetLineDim(obj.X + obj.Width, obj.Y, obj.X, obj.Y, 20, ksLineDimensionOrientationEnum.ksLinDHorizontal);
                                    SetConstrainttDim(Dim, obj, 1, 0, variable);

                                    Dim = SetLineDim(obj.X, obj.Y, obj.X, obj.Y + obj.Height, 20, ksLineDimensionOrientationEnum.ksLinDVertical);
                                    SetConstrainttDim(Dim, obj, 0, 1, variable);
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
                                if (this.doc.D5.ksMakeEncloseContours(0, obj.X1, obj.Y1) > 0)
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

                }

                //Привязка размера
             
            }
        }

        public void SetConstrainttDim(object dim, object obj, int Index1, int Index2, bool variable)
        {
            IDrawingObject1 Dim1 = (IDrawingObject1)dim;

            //Накладываем на объект ограничение совпадение точек 1
            IParametriticConstraint FixPoint1 = Dim1.NewConstraint();
            if (FixPoint1 != null)
            {
                FixPoint1.Comment = "FixPoint1";
                FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                FixPoint1.Partner = obj;
                FixPoint1.Index = 0;
                FixPoint1.PartnerIndex = Index1;
                if (FixPoint1.Create() == false) FixPoint1.Delete();
            }

            IParametriticConstraint FixPoint2 = Dim1.NewConstraint();
            if (FixPoint2 != null)
            {
                FixPoint2.Comment = "FixPoint2";
                FixPoint2.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                FixPoint2.Partner = obj;
                FixPoint2.Index = 1;
                FixPoint2.PartnerIndex = Index2;
                if (FixPoint2.Create() == false) FixPoint2.Delete();
            }

            if (variable)
            {
                //Накладываем на объект ограничение "Фиксированный размер"
                IParametriticConstraint FixetDim = Dim1.NewConstraint();
                if (FixetDim != null)
                {
                    FixetDim.Comment = "FixConstraint";
                    FixetDim.ConstraintType = ksConstraintTypeEnum.ksCFixedDim;
                    if (FixetDim.Create() == false) FixetDim.Delete();
                }


                //Накладываем на объект ограничение "Размер с переменной"
                IParametriticConstraint WithVariable = Dim1.NewConstraint();
                if (WithVariable != null)
                {
                    WithVariable.Comment = "WithVariable";
                    WithVariable.ConstraintType = ksConstraintTypeEnum.ksCDimWithVariable;
                    if (WithVariable.Create() == false) WithVariable.Delete();

                }
            }
        }

        //Сплитуем линию
        public async void SplitLine(ILineSegment lineSegment, double nSides, bool dellBaseSim)
        {
            this.doc.LockedLayerAsync(88, true);

            if (lineSegment != null)
            {
                IDrawingObject1 Constrait = (IDrawingObject1)lineSegment;

                List<object> Objs = new List<object>();//лист для линий.
                object[] segmentsArr = new object[2];
                if (Constrait.Constraints != null)
                {
                    foreach (IParametriticConstraint constraint in Constrait.Constraints)
                    {
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints)
                        {
                            if (constraint.Partner is Array array)
                            {
                                foreach (IDrawingObject drawingObject1 in array)
                                {
                                    if (drawingObject1.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                                    {
                                        ILineSegment lineSegments1 = (ILineSegment)drawingObject1;
                                        segmentsArr[constraint.Index] = lineSegments1;
                                    }
                                }
                            }
                            else
                            {
                                //если только один объект то пока не нужно
                            }
                        }
                    }
                }

                #region Создание линий
                ILineSegments segments = this.doc.GetDrawingContainer().LineSegments;
                if (segmentsArr[0] != null)
                    Objs.Add(segmentsArr[0]); //добавляем в список предыдущий

                //Начальная конечная точка
                System.Windows.Point start = new System.Windows.Point(lineSegment.X1, lineSegment.Y1);
                System.Windows.Point end = new System.Windows.Point(lineSegment.X2, lineSegment.Y2);

                //удаляем размер
                if (dellBaseSim)
                {
                    foreach (IParametriticConstraint constraint in Constrait.Constraints)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints)
                            if (constraint.Partner != null)
                                foreach (object TempObj in constraint.Partner)
                                    foreach (IParametriticConstraint constraint2 in Constrait.Constraints)
                                        if (constraint2.ConstraintType == ksConstraintTypeEnum.ksCMergePoints && constraint2.Index != constraint.Index)
                                            foreach (object TempObj2 in constraint2.Partner)
                                                if (TempObj == TempObj2)
                                                {
                                                    IDrawingObject drawingObject1 = (IDrawingObject)TempObj;
                                                    if (drawingObject1.DrawingObjectType == DrawingObjectTypeEnum.ksDrLDimension)
                                                        drawingObject1.Delete();
                                                }
                }

                //Удаляем линию
                lineSegment.Delete();


                ILineSegments lineSegments = this.doc.GetDrawingContainer().LineSegments;
                Point LastPoint = new Point(0, 0);

                for (int i = 0; i < nSides; i++)
                {
                    var tasks = new List<Task<object>>();
                    KmpsAppl.someFlag = false;
                    ILineSegment line = lineSegments.Add();
                    line.X1 = start.X + ((end.X - start.X) / nSides * i);
                    line.Y1 = start.Y + ((end.Y - start.Y) / nSides * i);
                    line.X2 = start.X + ((end.X - start.X) / nSides * (i + 1));
                    line.Y2 = start.Y + ((end.Y - start.Y) / nSides * (i + 1));
                    if (line.Update())
                    {
                        tasks.Add(Task<object>.Run(() =>
                        {
                            object point = CheckOrMakePoint(line, line.X1, line.Y1, 0, true);
                            return point; //поинт 1
                            }));
                        tasks.Add(Task<object>.Run(() =>
                        {
                            object point = CheckOrMakePoint(line, line.X2, line.Y2, 1, true);
                            return point; //поинт 2
                            }));
                        tasks.Add(Task<object>.Run(() =>
                        {
                            object obj = SetLineDim(line.X1, line.Y1, line.X2, line.Y2, 20);
                            return obj; //Размер 1
                            }));
                    }
                    Task.WhenAll(tasks).Wait();


                    ILineDimension lineDimension1 = (ILineDimension)tasks[2].Result;
                    lineDimension1.Update();

                    await Task.Run(() =>
                    {
                        this.SetConstrainttDim(lineDimension1, (object)line, 0, 1, true);
                        //this.doc.GetSelectContainer().Select(line);
                    });


                    Objs.Add((object)line); //Добавляем в писок линию

                    KmpsAppl.someFlag = true;


                }
                if (segmentsArr[1] != null)
                    Objs.Add(segmentsArr[1]); //добавляем в список следующую

                //Связываем линии
                for (int i = 0; i < Objs.Count - 1; i++)
                {
                    SetConstraintMergePoint(Objs[i], Objs[i + 1], 1, 0);
                }

                #endregion

            }

            this.doc.LockedLayerAsync(88, false);
            //Расселектируем
            this.doc.GetChooseContainer().UnchooseAll();
        }

        /// <summary>
        /// Make rectangle in work zone
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="Width"></param>
        /// <param name="Height"></param>

     
        /// <summary>
        /// Get selection object
        /// </summary>
        /// <returns></returns>
        private object SelectObject()
        {
            RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
            double x = 0, y = 0;

            int j = 1;
            while (j != 0)
            {
                info.Init();
                info.SetCursorText("Выберите объект");

                j = this.doc.D5.ksCursor(info, ref x, ref y, 0);
                if (j != 0)
                {
                    reference RefSelectedObj = this.doc.D5.ksFindObj(x, y, 100);
                    if (this.doc.D5.ksExistObj(RefSelectedObj) > 0)
                    {
                        SelectObj = (object)KmpsAppl.KompasAPI.TransferReference(RefSelectedObj, this.doc.D5.reference);
                        return SelectObj;
                    }
                }
            }
            return null;
        }

        //Привязка точки объектов
        public static bool SetConstraintMergePoint(object Obj1, object Obj2, int Index1, int Index2)
        {
            IDrawingObject1 obj1_1 = (IDrawingObject1)Obj1;

            //Накладываем на объект ограничение совпадение точек 1
            ParametriticConstraint FixPoint1 = obj1_1.NewConstraint();
            if (FixPoint1 != null)
            {
                FixPoint1.Comment = "FixPoint1";
                FixPoint1.ConstraintType = ksConstraintTypeEnum.ksCMergePoints;
                FixPoint1.Partner = Obj2;
                FixPoint1.Index = Index1;
                FixPoint1.PartnerIndex = Index2;
                if (FixPoint1.Create() == true)
                {
                    return true;
                }
                else
                {
                    FixPoint1.Delete();
                    return false;
                }
            }
            return false;

        }

        //удаляет ограничение согласно типу
        public static void RemoveConstraint(ksConstraintTypeEnum ksConstraintType, object Obj1, object Obj2 = null, int Index1 = -1)
        {
            IDrawingObject1 obj1_1 = (IDrawingObject1)Obj1;
            if (obj1_1.Constraints != null)
                foreach (IParametriticConstraint constraint in obj1_1.Constraints)
                        if (constraint.ConstraintType == ksConstraintType)
                            if ((constraint.Index == Index1) || (constraint.Index == -1))
                                if (Obj2 == null)
                                    constraint.Delete();
                                else
                                    if (constraint.Partner != null)
                                    foreach (object obj in constraint.Partner)
                                        if (obj == Obj2)
                                            constraint.Delete();
        }

        public static Tuple<double, double> GetIContourPerimentr (IContour contour)
        {
            if (contour != null)
            {
               double TLine = 0, TCurve = 0;

               for (int i = 0; i < contour.Count; i++)
                {
                    IContourSegment pDrawObj = (IContourSegment)contour.Segment[i];
                    // Получить тип объекта
                    // В зависимости от типа вывести сообщение для данного типа объектов
                    if (pDrawObj is IContourLineSegment contourLineSegment)
                    {
                        TLine += contourLineSegment.Length;
                    }
                    else
                    {
                        ICurve2D curve2d = pDrawObj.Curve2D;
                        TCurve += curve2d.Length;
                    }
                }
                return new Tuple<double, double>(TLine, TCurve);
            }
            return new Tuple<double, double>(0, 0);
        }

        /// <summary>
        /// Return intersection between two contour
        /// </summary>
        /// <param name="contour1"></param>
        /// <param name="contour2"></param>
        /// <returns>
        /// Item1 - Line perimeter, Item2 - Curve perimeter
        /// </returns>
        public static Tuple<double, double> IntersectionTwoIContour(IContour contour1, IContour contour2)
        {
            double LineInt = 0, CurveInt = 0;

            for (int j = 0; j < contour1.Count; j++)
            {
                IContourSegment segment = (IContourSegment)contour1.Segment[j];

                for (int k = 0; k < contour2.Count; k++)
                {
                    IContourSegment segment2 = (IContourSegment)contour2.Segment[k];

                    //Получаем крайние точки пересечения
                    if (segment.Curve2D != null)
                    {
                        double[] intersecArr = segment.Curve2D.Intersect(segment2.Curve2D);
                        if (intersecArr != null)
                        {
                            if (intersecArr.Length > 3)
                            {
                                //Узнаем длинну
                                double lenth = segment.Curve2D.GetDistancePointPoint(intersecArr[0], intersecArr[1], intersecArr[2], intersecArr[3]);
                                if (segment.SegmentType == ksContourSegmentEnum.ksCSLineSeg)
                                {
                                    LineInt += lenth;
                                }
                                else
                                {
                                    CurveInt += lenth;
                                }

                                //Если она больше, то нет смысла сравнивать дальше.
                                if (lenth >= segment.Curve2D.Length)
                                    break;
                            }
                        }
                    }
                }
            }

            return new Tuple<double, double>(LineInt, CurveInt);
        }

        public async Task SetVariableToDim(bool dell, ksDimensionTextBracketsEnum brackets, double VariableValue = 0)
        {
            dynamic objects = this.doc.GiveSelectOrChooseObj();

            if (objects is Array array)
            {
                //Перебираем объекты. Определяем что за объект внутри
                foreach (object obj in array)
                    await Task.Run(() =>
                    {
                        if (obj != null)
                            SetParamToObj(obj, dell, brackets, VariableValue);
                    });
            }
            else if (objects != null)
            {
                SetParamToObj((object)objects, dell, brackets, VariableValue);
            }

        }


        public void SetParamToObj(object objDim, bool dell, ksDimensionTextBracketsEnum bracketsEnum, double VariableValue = 0)
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
                                    (constraint.ConstraintType == ksConstraintTypeEnum.ksCFixedDim)) 
                                    constraint.Delete();
                    }
                    else
                    {
                        if (draw1.Constraints != null)
                            foreach (IParametriticConstraint constraint in draw1.Constraints)
                                if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                                {
                                    if (VariableValue == 0)
                                    {
                                        dimensionText.Brackets = bracketsEnum;
                                        lineDimension.Update();
                                        return;
                                    }
                                    else
                                    {
                                        SetValToVariableDim(draw1, VariableValue);
                                    }
                                }


                        //Накладываем на объект ограничение "Фиксированный размер"
                        IParametriticConstraint FixetDim = draw1.NewConstraint();
                        if (FixetDim != null)
                        {
                            FixetDim.Comment = "FixConstraint";
                            FixetDim.ConstraintType = ksConstraintTypeEnum.ksCFixedDim;
                            if (FixetDim.Create() == false)
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
                            if (WithVariable.Create() == false)
                            {
                                bracketsEnum = ksDimensionTextBracketsEnum.ksDimBrackets;
                                WithVariable.Delete();
                            }
                        }
                        if (VariableValue != 0) SetValToVariableDim(draw1, VariableValue);

                    }
                    dimensionText.Brackets = bracketsEnum;
                    lineDimension.Update();
                }
        }

        //Меняет параметр в размере
        public void SetValToVariableDim(IDrawingObject1 drawing1, double VariableValue)
        {
            if (drawing1.Constraints != null)
            {
                foreach (IParametriticConstraint constraint in drawing1.Constraints)
                    if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                        if (this.doc.Var.IsVariableNameValid(constraint.Variable) == false)
                        {
                            IVariable7 variable7 = this.doc.Var.Variable(constraint.Variable, string.Empty, false);
                            if (variable7 != null)
                                variable7.Value = VariableValue;
                            this.doc.D71.UpdateVariables();
                        }
            }
        }

        //Возвращает переменную
        public IVariable7 GetObjectVariable(IDrawingObject1 drawing1)
        {
            if (drawing1.Constraints != null)
                foreach (IParametriticConstraint constraint in drawing1.Constraints)
                    if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                        return this.doc.Var.Variable(constraint.Variable, string.Empty);
            return null;
        }

        //Соединяет две линии
        public void ConnectLineToLine()
        {
            this.doc.GetChooseContainer().UnchooseAll();

            object obj;

            if (this.doc.GetSelectContainer().SelectedObjects is Array array)
            {
                obj = array.GetValue(array.Length - 1);
            }
            else 
            {
                obj = this.doc.GetSelectContainer().SelectedObjects;
            }

            if (obj != null)
            {
                this.doc.LockedLayerAsync(88, true);
                KmpsAppl.someFlag = false;

                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
                //Выделяем объект которому будем привязываться
                info.Init();
                info.SetCursorText("Выберите объект привязки");
                SelectObj = SelectObject();
                LightDotOnObj(SelectObj, true);
                double x = 0, y = 0;
                this.doc.D5.ksCursor(info, ref x, ref y, phtm);
                int indexSelectObj = NumberPointNear(SelectObj, x, y);
                LightDotOnObj(SelectObj, false);
                SetConstraintMergePoint(obj, SelectObj, GetMergeIndex(obj), indexSelectObj);
            }

            this.doc.LockedLayerAsync(88, false);
            KmpsAppl.someFlag = true;

            int GetMergeIndex(object fObj)
            {
                List<int> BreakIndex = new List<int>();
                IDrawingObject1 drawingObject1 = (IDrawingObject1)fObj;
                if (drawingObject1.Constraints != null)
                    foreach (ParametriticConstraint constraint in drawingObject1.Constraints)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints)
                            if (BreakIndex.IndexOf(constraint.Index) == -1)
                                foreach (IDrawingObject drawingObject in constraint.Partner)
                                    if (drawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                                        if (BreakIndex.IndexOf(constraint.Index) == -1)
                                            BreakIndex.Add(constraint.Index);
                if (BreakIndex.IndexOf(0) == -1) return 0;
                if (BreakIndex.IndexOf(1) == -1) return 1;

                return -1;
            }
        }
        //Инвертирует размеры на координатном раскрое
        public void InvertPointCoord(bool xinvert)
        {
            KmpsAppl.someFlag = false;
            ISelectionManager selection = this.doc.GetSelectContainer();

                if (selection.SelectedObjects is Array array)
                {
                    foreach (object obj in array)
                    {
                        invert(obj);
                    }
                }
                else if (selection.SelectedObjects != null)
                {
                    invert(selection.SelectedObjects);
                }

            KmpsAppl.someFlag = true;

            void invert (object obj)
            {
                if (obj is IBaseLeader bLeader)
                {
                    SetPoint(bLeader, xinvert);
                }
                
                void SetPoint(IBaseLeader baseLeader, bool invertX)
                {
                    IAttribute attribute = (IAttribute)KmpsAppl.KompasAPI.TransferReference(this.doc.Attribute.GiveObjAttr(baseLeader.Reference), this.doc.D5.reference);

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

                        double cx = x + sizeX / 2;
                        double cy = y + sizeY / 2;

                        IBranchs branchs = (IBranchs)baseLeader;

                        double objX = branchs.BranchX[0];
                        double objY = branchs.BranchY[0];


                        ILeader leader = (ILeader)baseLeader;
                        IText TextOnShelf = leader.TextOnShelf;
                        ITextLine textLineOn = TextOnShelf.TextLine[0];
                        ITextItem itemOnShelf = textLineOn.TextItem[1];
                        ITextFont fontOnShelf = (ITextFont)itemOnShelf;

                        IText TextUnderShelf = leader.TextUnderShelf;
                        ITextLine BrachText = TextUnderShelf.Add();

                        int iT = 0;
                        //Проверяем X
                        if (FindOrd(textLineOn, "x"))
                        {
                            if ((Math.Round(objX, 0) != Math.Round(x, 0)) && (Math.Round(objX, 0) != Math.Round(dx, 0)))
                            {
                                ITextItem textItem = BrachText.Add();
                                textItem.Str = (objX >= cx && invertX ?  Math.Round(x - objX, 1) : Math.Round(dx - objX, 1)).ToString();
                                textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                ITextFont font = (ITextFont)textItem;
                                font.Height = fontOnShelf.Height;
                                textItem.Update();
                                iT++;
                                ITextItem textItemUP = BrachText.Add();
                                textItemUP.Str = "x";
                                textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                ITextFont fontUP = (ITextFont)textItem;
                                font.Height = fontOnShelf.Height;
                                textItemUP.Update();
                                iT++;
                            }
                        }
                        else
                        {
                            if (leader.ShelfDirection == ksShelfDirectionEnum.ksLSLeft) leader.ShelfDirection = ksShelfDirectionEnum.ksLSRight;
                            else leader.ShelfDirection = ksShelfDirectionEnum.ksLSLeft;
                        }
                        //Проверяем Y
                        if (FindOrd(textLineOn, "y"))
                        {
                            if ((Math.Round(objY, 0) != Math.Round(y, 0)) && (Math.Round(objY, 0) != Math.Round(dy, 0)))
                            {
                                ITextItem textItem = BrachText.Add();
                                textItem.Str = (objY >= cy && !invertX ? Math.Round(y - objY, 1) : Math.Round(dy - objY, 1)).ToString();
                                textItem.ItemType = ksTextItemEnum.ksTItDeviationEnd;
                                ITextFont font = (ITextFont)textItem;
                                font.Height = fontOnShelf.Height;
                                textItem.Update();
                                iT++;

                                ITextItem textItemUP = BrachText.Add();
                                textItemUP.Str = "y";
                                textItemUP.ItemType = ksTextItemEnum.ksTItUpperDeviation;
                                ITextFont fontUP = (ITextFont)textItem;
                                font.Height = fontOnShelf.Height;
                                textItemUP.Update();
                                iT++;
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

        //Ищет индекс точки ближайшей к координатам
        private int NumberPointNear(object Obj, double xm, double my)
        {
            IDrawingObject drawingObject = (IDrawingObject)Obj;
            int pointIndex = -1;

            if (drawingObject != null)
                switch (drawingObject.DrawingObjectType)
                {
                    case DrawingObjectTypeEnum.ksDrLineSeg:
                        ILineSegment LineObj = (ILineSegment)Obj;
                        if (KmpsAppl.Mat.ksDistancePntPnt(LineObj.X1, LineObj.Y1, xm, my) > KmpsAppl.Mat.ksDistancePntPnt(LineObj.X2, LineObj.Y2, xm, my))
                            return 1;
                        else
                            return 0;

                }

            LightDotOnObj(SelectObj, true, pointIndex);
            return pointIndex;
        }

        private void LightDotOnObj(object obj, bool on, int indexOn = -1)
        {
            IDrawingObject1 drawingObject1 = (IDrawingObject1)obj;

            if (drawingObject1 != null)
            {
                foreach (ParametriticConstraint constraint in drawingObject1.Constraints)
                {
                    if ((constraint.ConstraintType == ksConstraintTypeEnum.ksCMergePoints) && (constraint.Partner != null))
                    {
                        foreach (IDrawingObject drawingObject in constraint.Partner)
                        {
                            if (drawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrPoint)
                            {
                                IPoint point = (IPoint)drawingObject;
                                if (on && (constraint.Index == indexOn || indexOn == -1))
                                {
                                    this.doc.GetChooseContainer().Choose(point);
                                    point.Style = 4;
                                }
                                else
                                {
                                    this.doc.GetChooseContainer().Unchoose(point);
                                    point.Style = 1;
                                }

                                point.Update();
                            }
                        }
                    }
                }
            }
        }

        //Рисует размер между двумя объектами
        public void ObjectsToObjectDim()
        {
            this.doc.LockedLayerAsync(88, true);
            KmpsAppl.someFlag = false;
            //Забираем выделенные объекты
            ISelectionManager selection = this.doc.GetSelectContainer();

            RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
            double x = 0, y = 0;

            int indexSelectObj = -1;
            //Выделяем объект которому будем привзываться
            info.Init();
            info.SetCursorText("Выберите объект привязки");
            SelectObj = SelectObject();

            if (SelectObj != null)
            {
                while (indexSelectObj < 0)
                {
                    //Селектируем
                    this.doc.GetChooseContainer().Choose(SelectObj);

                    SelectPointFlag = true;

                    LightDotOnObj(SelectObj, true);
                    //Выбираем
                    info.Init();
                    info.SetCursorText("Выберите точку");
                    if (this.doc.D5.ksCursor(info, ref x, ref y, phtm) != 0)
                    {

                        indexSelectObj = NumberPointNear(SelectObj, x, y);

                        //Идем по объектам и выбираем точки привязки
                        if (selection.SelectedObjects is Array array) 
                        { 
                            foreach (IDrawingObject drawingObject in array)
                            {
                                MakeObjToObjDim(drawingObject);
                            }
                        }
                        else if (selection.SelectedObjects != null)
                        {
                            MakeObjToObjDim((IDrawingObject)selection.SelectedObjects);

                        }

                        LightDotOnObj(SelectObj, false);
                    }
                    else
                        break;

                    void MakeObjToObjDim(IDrawingObject IdObj)
                    {
                        double x2 = 0, y2 = 0;
                        //Селектируем
                        this.doc.GetChooseContainer().Choose(IdObj);
                        LightDotOnObj(IdObj, true);
                        SelectPointFlag = true;
                        //Выбираем
                        info.SetCursorText("Выберите точку");
                        this.doc.D5.ksCursor(info, ref x2, ref y2, phtm);
                       int indexDrawObj = NumberPointNear(IdObj, x2, y2);
                        this.doc.GetChooseContainer().Unchoose(IdObj);
                        SetConstrainttDim(SetLineDim(x, y, x2, y2, 0), IdObj, 0,indexDrawObj, false);
                        LightDotOnObj(IdObj, false);
                    }
                }
                this.doc.GetChooseContainer().UnchooseAll();
            }
            this.doc.LockedLayerAsync(88, false);
            KmpsAppl.someFlag = true;
        }

        //Создает или обновляет линию
        public async void MakeLine(double LenthValue, double Angle, bool Adding, bool single)
        {
            this.WorkFlag = true;
            if (Adding)
            {
                //проверяем есть ли объекты
                if (lastLine.Count > 0)
                {
                    if (lastLine[0] != null)
                        if (this.doc.D5.ksExistObj(lastLine[0].Reference) == 0) lastLine.Clear();
                    if (lastLine.Count > 0)
                        if (this.doc.D5.ksExistObj(lastLine.Last().Reference) == 0) lastLine.Remove(lastLine.Last());
                    if (firstDim != null)
                        if (this.doc.D5.ksExistObj(firstDim.Reference) == 0) firstDim = null;
                }

                ILineSegments lineSegments = this.doc.GetDrawingContainer().LineSegments;

                if (this.doc.GetSelectContainer().SelectedObjects != null && !(this.doc.GetSelectContainer().SelectedObjects is object[]))
                {       //Если выбранный объект линия то она будет первым объектом
                    if (((IDrawingObject)this.doc.GetSelectContainer().SelectedObjects).DrawingObjectType == DrawingObjectTypeEnum.ksDrAnnLineSeg)
                    {
                        lastLine[0] = (ILineSegment)this.doc.GetSelectContainer().SelectedObjects;
                        this.doc.GetSelectContainer().UnselectAll();
                    }
                }


                if (lastLine.Count == 0) //Если объект первой линии пустой то мы ему находим 
                    if (lineSegments.Count > 0)
                        lastLine.Add(lineSegments.LineSegment[0]);

                if (lastLine.Count > 1)
                {
                    //Отвязываем точку последнего и первого
                    RemoveConstraint(ksConstraintTypeEnum.ksCMergePoints, lastLine[0], lastLine.Last(), 0);
                    RemoveConstraint(ksConstraintTypeEnum.ksCMergePoints, firstDim, lastLine.Last(), 0);
                    //Привязываем обратно размер к линии.
                    SetConstraintMergePoint(lastLine[0], firstDim, 0, 0);


                    double step = 360 / lastLine.Count;

                    if (lastLine.Count > 3)
                        for (int j = 2; j < lastLine.Count; j++)
                        {
                            lastLine[j].Angle = lastLine[j - 1].Angle - step;
                            lastLine[j].Update();
                        }
                }

                //Добавляем линию
                ILineSegment lineSegment = lineSegments.Add();

                if (lastLine.Count > 1) //Если есть последний объект
                {
                    lineSegment.X1 = lastLine.Last().X2;
                    lineSegment.Y1 = lastLine.Last().Y2;
                    lineSegment.X2 = lastLine[0].X1;
                    lineSegment.Y2 = lastLine[0].Y1;

                    if (lastLine.Last() == lastLine[0])
                    {
                        lineSegment.X2 = LenthValue;
                        lineSegment.Y2 = lastLine[0].Y2;
                    }
                }
                else //Если нет последнего объекта.
                {
                    lineSegment.Length = LenthValue;
                    lineSegment.X1 = 0;
                    lineSegment.Y1 = 0;
                    lineSegment.Angle = 90;
                    lastLine.Add(lineSegment);
                }
                if (lineSegment.Update())
                {
                    var tasks = new List<Task<object>>
                    {
                        Task<object>.Run(() =>
                        {
                            return (object)CheckOrMakePoint(lineSegment, lineSegment.X1, lineSegment.Y1, 0, true); //поинт 1
                    }),
                        Task<object>.Run(() =>
                        {
                            return (object)CheckOrMakePoint(lineSegment, lineSegment.X2, lineSegment.Y2, 1, true); //поинт 2
                    }),
                        Task<object>.Run(() =>
                        {
                            return (object)SetLineDim(lineSegment.X1, lineSegment.Y1, lineSegment.X2, lineSegment.Y2, 20); ; //Размер 1
                    })
                    };


                    Task.WhenAll(tasks).Wait();


                    if (lastLine.Count > 2)
                    {
                        await Task.Run(() =>
                        {
                            //Привязываем первую точку
                            SetConstraintMergePoint(lineSegment, lastLine.Last(), 0, 1);
                            //Привязываем вторую точку
                            SetConstraintMergePoint(lineSegment, lastLine[0], 1, 0);
                        });

                    }
                    else if (lastLine.Count > 1)
                    {
                        await Task.Run(() =>
                        {
                            SetConstraintMergePoint(lineSegment, lastLine[0], 0, 1);
                        });
                    }


                    //Образмериваем
                    ILineDimension lineDimension = (ILineDimension)tasks[2].Result;
                    if (lastLine.Count == 1)
                    {
                        await Task.Run(() =>
                        {
                            firstDim = lineDimension;
                            FixFirstLine(lineDimension);
                        });
                    }

                    await Task.Run(() =>
                    {
                        //Связываем размер с линией
                        SetConstrainttDim(lineDimension, (object)lineSegment, 0, 1, true);

                        //Фиксируем и даем параметр
                        SetValToVariableDim((IDrawingObject1)lineDimension, LenthValue);

                        RemoveConstraint(ksConstraintTypeEnum.ksCHorizontal, lineDimension);
                    });

                    //Теперь это последний объект
                    lastLine.Add(lineSegment);
                }
            }

            this.WorkFlag = false;

            KmpsAppl.ZoomAll();

            //апдейт переменной
           

            void FixFirstLine(object Obj)
            {
                IDrawingObject1 line1_1 = (IDrawingObject1)Obj;

                if (line1_1 != null)
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
                    IParametriticConstraint FixLine2 = line1_1.NewConstraint();
                    if (FixLine2 != null)
                    {
                        FixLine2.Comment = "BaseLine";
                        FixLine2.ConstraintType = ksConstraintTypeEnum.ksCFixedPoint;
                        FixLine2.Index = 1;
                        FixLine2.Create();
                    }
                }
            }


        }

        private IPoint CheckOrMakePoint(object obj, double X, double Y, int index, bool upd)
        {
            IPoints points = this.doc.GetDrawingContainer().Points;

            foreach (IPoint point in points)
                if (point.X == X && point.Y == Y)
                    return null;

            IPoint newpoint = points.Add();
            newpoint.X = X;
            newpoint.Y = Y;
            newpoint.Style = (int)ksAnnotationSymbolEnum.ksDotPoint;
            newpoint.LayerNumber = 88;
            if (upd) newpoint.Update();

            SetConstraintMergePoint(obj, newpoint, index, 0);

            return newpoint;
        }

        //Тупой угол
        public void ObtuseAngle()
        {
            if (this.doc != null)
            {
                this.doc.LockedLayerAsync(88, true);

                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
                double x = 0, y = 0;

                List<ILineSegment> segments = new List<ILineSegment>(); //лист для сегментов
                this.doc.GetChooseContainer().UnchooseAll(); //убираем все выделения
                ISelectionManager selection = this.doc.GetSelectContainer(); //Получаем выбранные объекты

                if (selection.SelectedObjects is Array array)
                {
                    foreach (IDrawingObject obj in array)
                        if (obj.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                            segments.Add((ILineSegment)obj); //Забираем выбранные линии
                }
                else if (selection.SelectedObjects != null)
                {
                    object pObj = selection.SelectedObjects;
                    IDrawingObject obj = (IDrawingObject)pObj;
                    if (obj != null)
                        if (obj.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                            segments.Add((ILineSegment)obj); //Если линия одна
                }

                reference RefSelectedObj = -1; int j = -1;
                while ((segments.Count < 2) && (j != 0)) //Добираем объекты до двух если их количество меньше
                {
                    info.Init();
                    info.SetCursorText("Выберите объект");

                    j = this.doc.D5.ksCursor(info, ref x, ref y, 0);
                    RefSelectedObj = this.doc.D5.ksFindObj(x, y, 100);
                    if (this.doc.D5.ksExistObj(RefSelectedObj) > 0)
                    {
                        SelectObj = (object)KmpsAppl.KompasAPI.TransferReference(RefSelectedObj, this.doc.D5.reference);
                        IDrawingObject obj = (IDrawingObject)SelectObj;
                        if (obj.DrawingObjectType == DrawingObjectTypeEnum.ksDrLineSeg)
                        {
                            segments.Add((ILineSegment)obj);
                            this.doc.GetChooseContainer().Choose(SelectObj);
                        }
                    }
                }
                if (segments.Count < 2) return; //если все еще не набрали, то заканчиваем
                else
                {
                    for (int i = 1; i < segments.Count; i++)
                    {
                        if (!CheckSize(segments[i - 1], segments[i]))
                        {
                            SetConstraintoADim(SetAngleDim(segments[i - 1], segments[i]), ksDimensionTextBracketsEnum.ksDimBracketsOff);
                        }
                    }
                }

                this.doc.LockedLayerAsync(88, false);
                this.doc.GetChooseContainer().UnchooseAll();
            }

            bool CheckSize(ILineSegment segment1, ILineSegment segment2)
            {
                IDrawingObject1 drawingObject1 = (IDrawingObject1)segment1;
                IDrawingObject1 drawingObject2 = (IDrawingObject1)segment2;

                if (drawingObject1.Constraints != null)
                    foreach (IParametriticConstraint constraint in drawingObject1.Constraints)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                            if (constraint.Partner != null)
                                foreach (IDrawingObject drawingObject in constraint.Partner)
                                    if (drawingObject.DrawingObjectType == DrawingObjectTypeEnum.ksDrADimension)
                                        return true;

                return false;
            }

            void SetConstraintoADim(IAngleDimension angleDimension, ksDimensionTextBracketsEnum bracketsEnum)
            {
                int Angle = 90;

                IDrawingObject objDim = (IDrawingObject)angleDimension;
                IDimensionText dimensionText = (IDimensionText)angleDimension;
                dimensionText.Brackets = bracketsEnum;
                angleDimension.Update();

                IDrawingObject1 Dim1 = (IDrawingObject1)objDim;

                //Фиксируем точку на координатах или объекте
                //Накладываем на объект ограничение "Фиксированный размер"
                IParametriticConstraint FixetDim = Dim1.NewConstraint();
                if (FixetDim != null)
                {
                    FixetDim.Comment = "FixConstraint";
                    FixetDim.ConstraintType = ksConstraintTypeEnum.ksCFixedDim;
                    if (!FixetDim.Create())
                    {
                        bracketsEnum = ksDimensionTextBracketsEnum.ksDimBrackets;
                        FixetDim.Delete();
                    }
                    else
                    {
                        //Накладываем на объект ограничение "Размер с переменной"
                        IParametriticConstraint WithVariable = Dim1.NewConstraint();
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


                        IParametriticConstraint Associate = Dim1.NewConstraint();
                        if (Associate != null)
                        {

                            ksDynamicArray arrayCurve = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(ldefin2d.ARRAYPARAMTABLE_OBJ);
                            arrayCurve.ksAddArrayItem(0, angleDimension.BaseObject1);
                            arrayCurve.ksAddArrayItem(1, angleDimension.BaseObject2);

                            Associate.Comment = "Associate";
                            Associate.ConstraintType = ksConstraintTypeEnum.ksCAssociation;
                            Associate.Partner = arrayCurve;
                            if (!Associate.Create())
                                Associate.Delete();
                        }
                    }
                }
                SetValToVariableDim(Dim1, Angle);
            }

        }



    }

}
