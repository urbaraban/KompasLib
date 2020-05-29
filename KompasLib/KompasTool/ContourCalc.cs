using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Tools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using reference = System.Int32;

namespace KompasLib.KompasTool
{
    public static class ContourCalc
    {
        public static void GetContour(int _index, ItemCollection CeilingItems, double width,
            bool Usadka, bool Mashtab, 
            double Comp, int stg_crd_dopusk,
            bool cursor = true, reference refContour = 0)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                #region получение контура
                double x = 0, y = 0;
                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);
                string index = CeilingItems[_index].ToString();

                //Ищем или находим макрообъект по индексу потолка
                IMacroObject macroObject = KmpsAppl.Doc.Macro.FindCeilingMacro(index);
                if (macroObject == null) macroObject = KmpsAppl.Doc.Macro.MakeCeilingMacro(index);
                else if (cursor)
                {
                    KmpsAppl.Doc.Macro.RemoveCeilingMacro(index);
                    macroObject = KmpsAppl.Doc.Macro.MakeCeilingMacro(index);
                }

                if (refContour == 0)
                {
                    if (cursor)
                    {
                        KmpsAppl.Doc.D5.ksCursor(info, ref x, ref y, 0);
                        int refEncloseContours = KmpsAppl.Doc.D5.ksMakeEncloseContours(0, x, y);
                        if (refEncloseContours != 0)
                            KmpsAppl.Doc.Macro.AddCeilingMacro(refEncloseContours, index);  //Добавляем ksMakeEncloseContours
                        else
                            KmpsAppl.KompasAPI.ksMessage("Не найден замкнутый контур");

                        refContour = KmpsAppl.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                    }
                    else
                    {
                        refContour = KmpsAppl.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                    }
                }
                else
                {
                    KmpsAppl.Doc.Macro.RemoveCeilingMacro(index);
                    macroObject = KmpsAppl.Doc.Macro.MakeCeilingMacro(index);
                    KmpsAppl.Doc.Macro.AddCeilingMacro(refContour, index);
                    refContour = KmpsAppl.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                }
                #endregion
                if (refContour != 0)
                    CalcAll(macroObject, Usadka);  //Считает все
            }

            async void CalcAll(IMacroObject macroObject1, bool usadka)
            {
                string index = CeilingItems[_index].ToString();
                var tasks = new List<Task<bool>>();

                if (usadka)
                {
                    tasks.Add(Task<bool>.Run(() => KVariable.UpdateAsync("PerimetrU", 0, index)));
                    if (_index == 0) tasks.Add(Task<bool>.Run(() => KVariable.UpdateAsync("Angle", 0, string.Empty)));
                }
                else
                {
                    tasks.Add(Task<bool>.Run(() =>  KVariable.UpdateAsync("Perimetr", 0, index)));
                    tasks.Add(Task<bool>.Run(() =>  KVariable.UpdateAsync("LineP", 0, index)));
                    tasks.Add(Task<bool>.Run(() =>  KVariable.UpdateAsync("CurveP", 0, index)));
                    tasks.Add(Task<bool>.Run(() =>  KVariable.UpdateAsync("Shov", 0, index)));
                    tasks.Add(Task<bool>.Run(() =>  KVariable.UpdateAsync("cut", 0, index)));
                    if (_index == 0) tasks.Add(Task<bool>.Run(() => KVariable.UpdateAsync("Angle", 0, string.Empty)));
                }
                Task.WhenAll(tasks).Wait();
                tasks.Clear();

                ksInertiaParam inParam = (ksInertiaParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_InertiaParam);
                ksIterator Iterator1 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                Iterator1.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject1.Reference);
                double SqareMain = 0;
                int position = CeilingItems.IndexOf(index);
                int MainRef = 0;
                int Count = 0;
                //Перебираем кривые в поиске самой большой.
                reference refContour1 = Iterator1.ksMoveIterator("F");
                while (refContour1 != 0)
                {
                    KmpsAppl.Mat.ksCalcInertiaProperties(refContour1, inParam, ldefin2d.ST_MIX_M);
                    if (SqareMain < inParam.F)
                    {
                        SqareMain = inParam.F; //Выбираем самую большую
                        MainRef = refContour1;
                    }
                    Count++;
                    refContour1 = Iterator1.ksMoveIterator("N");
                }
                //
                //Получить габарит mainRef
                //
                MakeGabarit(MainRef, Mashtab);

                //
                //Начинаем перебор контуров со всем что есть
                //
                //Заходим в первый контур


                refContour1 = Iterator1.ksMoveIterator("F");
                while (refContour1 != 0)
                {

                    double lineTemp = 0, curveTemp = 0, shovTemp = 0, cut = 0, angleTemp = KVariable.Give("Angle", string.Empty);
                    IContour contour1 = KmpsAppl.Doc.Macro.GiveContour(refContour1);
                    if (refContour1 != MainRef)
                    {
                        KmpsAppl.Mat.ksCalcInertiaProperties(refContour1, inParam, 3);   //Если он не главный, то вычитаем его площадь из главного.
                        SqareMain -= inParam.F;                                 //Вычитаем из главной площади
                        perriContour(contour1, ref cut, ref cut, ref angleTemp);               //Считаем предполагаем вырез
                    }
                    else perriContour(contour1, ref lineTemp, ref curveTemp, ref angleTemp);
                    //Функция расчета периметра
                    void perriContour(IContour contour, ref double Line, ref double Curve, ref double Angle)
                    {
                        if (contour != null)
                        {
                            for (int i = 0; i < contour.Count; i++)
                            {
                                IContourSegment pDrawObj = (IContourSegment)contour.Segment[i];
                                // Получить тип объекта
                                // В зависимости от типа вывести сообщение для данного типа объектов
                                try
                                {
                                    IContourLineSegment contourLineSegment = (IContourLineSegment)pDrawObj;
                                    Line += contourLineSegment.Length;
                                    Angle += 1;
                                }
                                catch
                                {
                                    ICurve2D contourLineSegment = (ICurve2D)pDrawObj.Curve2D;
                                    Curve += contourLineSegment.Length;
                                    if (i > 0)
                                        if (contour.Segment[i - 1].Type == KompasAPIObjectTypeEnum.ksObjectContourLineSegment)
                                            Angle += 1;
                                }
                            }
                        }
                    }


                    KmpsAppl.ProgressBar.Start(0, CeilingItems.Count, "Считаем контур ", true);

                    for (int i = 0; i < CeilingItems.Count; i++)           //Стартуем проход по потолкам
                    {
                        KmpsAppl.ProgressBar.SetProgress(i, "Считаем контур " + i, false);
                        if (i == position)
                            continue;                            //Если это тот же потолок то выходим.
                        else
                        {
                            IMacroObject macroObject2 = KmpsAppl.Doc.Macro.FindCeilingMacro(CeilingItems[i].ToString());
                            if (macroObject2 != null)
                            {
                                ksIterator Iterator2 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                                Iterator2.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject2.Reference);

                                //Заходим во второй контур
                                reference refContour2 = Iterator2.ksMoveIterator("F");
                                while (refContour2 != 0)
                                {
                                    IContour contour2 = KmpsAppl.Doc.Macro.GiveContour(refContour2);
                                    ksDynamicArray arrayCurve = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(ldefin2d.POINT_ARR);
                                    if (KmpsAppl.Mat.ksIntersectCurvCurv(contour1.Reference, contour2.Reference, arrayCurve) == 1)
                                    {
                                        angleTemp -= 1;
                                        int step = (int)arrayCurve.ksGetArrayCount();
                                        ksMathPointParam[] points = new ksMathPointParam[step];

                                        for (int j = 0; j < step; j++)
                                        {
                                            points[j] = (ksMathPointParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
                                            arrayCurve.ksGetArrayItem(j, points[j]);
                                        }

                                        double lenthTemp = 0;

                                        //Сравниваем сегменты контура с другим контуром (по сегментно)
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
                                                        if (intersecArr.Length > 3)
                                                        {
                                                            //Узнаем длинну
                                                            double lenthTemp2 = segment.Curve2D.GetDistancePointPoint(intersecArr[0], intersecArr[1], intersecArr[2], intersecArr[3]);
                                                            lenthTemp += lenthTemp2;

                                                            //Если она больше, то нет смысла сравнивать дальше.
                                                            if (lenthTemp2 >= segment.Curve2D.Length)
                                                                break;
                                                        }
                                                }
                                            }
                                        }

                                        shovTemp += lenthTemp;

                                        arrayCurve.ksDeleteArray();
                                        if (i < position)
                                        {
                                            if (Math.Round(curveTemp - shovTemp, 2) < 0) { lineTemp += curveTemp - shovTemp; curveTemp = 0; }
                                            else curveTemp -= shovTemp;
                                            shovTemp = 0;
                                        }
                                    }
                                    refContour2 = Iterator2.ksMoveIterator("N"); //Двигаем итератор 2

                                    arrayCurve.ksDeleteArray();
                                }
                                Iterator2.ksDeleteIterator(); //Удаляем итератор 2
                            }
                        }
                    }
                    //получили шов и вычитаем его из предыдущих параметров периметра или длинны выреза.

                    shovTemp = Math.Round(shovTemp, 2);
                    curveTemp = Math.Round(curveTemp, 2);
                    lineTemp = Math.Round(lineTemp, 2);

                    if (refContour1 != MainRef) cut -= shovTemp;
                    else
                    {
                        if (curveTemp - shovTemp <= 0) { lineTemp += curveTemp - shovTemp; curveTemp = 0; }
                        else curveTemp -= shovTemp;
                    }



                    if (usadka)
                    {
                        tasks.Add(Task<object>.Run(() => KVariable.Add("PerimetrU", Math.Round((lineTemp + curveTemp) / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("Angle", angleTemp, string.Empty)));
                    }
                    else
                    {
                        tasks.Add(Task<object>.Run(() => KVariable.Add("Perimetr", Math.Round((lineTemp + curveTemp) / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("LineP", Math.Round(lineTemp / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("CurveP", Math.Round(curveTemp / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("Shov", Math.Round(shovTemp / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("cut", Math.Round(cut / (1000 / Comp), 2), index)));
                        tasks.Add(Task<object>.Run(() => KVariable.Add("Angle", angleTemp, string.Empty)));
                    }

                    Task.WhenAll(tasks).Wait();

                    refContour1 = Iterator1.ksMoveIterator("N"); //Двигаем итератор 1
                }
                Iterator1.ksDeleteIterator(); //Удаляем итератор 1 после полного перебора
                if (usadka) await KVariable.UpdateAsync("SqareU", Math.Round(SqareMain / (Math.Pow(10, 6) / Math.Pow(Comp, 2)), 2), index);
                else await KVariable.UpdateAsync("Sqare", Math.Round(SqareMain / (Math.Pow(10, 6) / Math.Pow(Comp, 2)), 2), index);

                KmpsAppl.ProgressBar.Stop("Закончили", true);

                async void MakeGabarit(reference objRef, bool garpun)
                {
                    ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                    ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                    if (KmpsAppl.Doc.D5.ksGetObjGabaritRect(objRef, spcGabarit) == 1)
                    {
                        ksMathPointParam mathBop = spcGabarit.GetpBot();
                        ksMathPointParam mathTop = spcGabarit.GetpTop();
                        double x = mathBop.x;
                        double y = mathBop.y;
                        double dx = mathTop.x;
                        double dy = mathTop.y;
                        double sizeX = Math.Round(Math.Abs(x - dx), 2);
                        double sizeY = Math.Round(Math.Abs(y - dy), 2);

                        await KVariable.UpdateAsync("Xcrd", x, index);
                        await KVariable.UpdateAsync("Ycrd", y, index);

                        if (!usadka)
                        {
                            await KVariable.UpdateAsync("XGabarit", sizeX, index);
                            await KVariable.UpdateAsync("YGabarit", sizeY, index);
                        }

                        if (usadka || !garpun)
                        {
                            IVariable7 lenth = KmpsDoc.D71.Variable[false, "lenth" + index];

                            recPar.Init();
                            recPar.x = x;
                            recPar.y = y;
                            double Dopusk = stg_crd_dopusk / 100 + 1;

                            if ((sizeX / width) <= 1 * Dopusk && (sizeX / width) > (sizeY / width) || (sizeY > width * Dopusk))
                            {
                                lenth.Value = sizeY;
                                if (y < dy) recPar.height = sizeY;
                                else recPar.height = -sizeY;
                                if (x < dx) recPar.width = width;
                                else recPar.width = -width;
                            }
                            else
                            {
                                lenth.Value = sizeX;
                                if (y < dy) recPar.height = width;
                                else recPar.height = -width;
                                if (x < dx) recPar.width = sizeX;
                                else recPar.width = -sizeX;
                            }

                            KmpsDoc.D71.UpdateVariables();
                        }
                    }
                }
            }

        }
    }
}
