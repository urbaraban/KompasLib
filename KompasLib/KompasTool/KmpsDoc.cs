using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace KompasLib.Tools
{
    public class KmpsDoc
    {
        private static ksDocument2D doc5;
        private IKompasDocument2D doc7;
        private static IKompasDocument2D1 doc71;
        private IKompasDocument1 doc1;
        private ksDocumentParam docPar;
        private IDocuments documents7;
        private SizeTool sizeTool;
        private KmpsMacro macro;
        private KmpsAttr attr;


        //Селектирован размер
        public event EventHandler<object> SelectDimenetion;

        public SizeTool ST
        {
            get => sizeTool;
        }

        public KmpsAttr Attribute
        {
            get => attr;
        }

        public KmpsDoc()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                this.documents7 = KmpsAppl.Appl.Documents;

                // Получаем интерфейс активного документа 2D в API7
                this.doc7 = (IKompasDocument2D)KmpsAppl.Appl.ActiveDocument;

                if (doc7 == null)
                    return;



                KmpsDoc.doc71 = (IKompasDocument2D1)KmpsAppl.Appl.ActiveDocument;
                this.doc1 = (IKompasDocument1)KmpsAppl.Appl.ActiveDocument;

                attr = new KmpsAttr();
                
                attr.FuncAttrType();

                // Получаем интерфейс активного документа 2D в API5
                doc5 = (ksDocument2D)KmpsAppl.KompasAPI.ActiveDocument2D();
                if (doc5 == null)
                    return;

                this.docPar = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
                if (docPar == null)
                    return;
                else this.docPar.type = (short)DocType.lt_DocFragment;

                this.macro = new KmpsMacro();
                this.sizeTool = new SizeTool(this);
            }
        }
    
        public IKompasDocument2D D7
        {
            get => doc7;
            set => doc7 = value;
        }

        public IKompasDocument1 D1
        {
            get => doc1;
            set => doc1 = value;
        }


        public static IKompasDocument2D1 D71
        {
            get => doc71;
        }

        public ksDocument2D D5
        {
            get => doc5;
            set => doc5 = value;
        }

        public IDocuments Documents7
        {
            get => documents7;
            set => documents7 = value;
        }

        public ksDocumentParam DocPar
        {
            get => docPar;
            set => docPar = value;
        }

        public KmpsMacro Macro
        {
            get => macro;
            set => macro = value;
        }

        //Получает только селектированные объекты
        public ISelectionManager GetSelectContainer()
        {
            if (doc71 != null)
            {
                // Получить менеджер видов и слоев
                ISelectionManager mng = doc71.SelectionManager;

                if (mng != null)
                    return mng;
            }
            return null;
        }

        //Получает только выделенные объекты
        public IChooseManager GetChooseContainer()
        {
            if (doc71 != null)
            {
                // Получить менеджер видов и слоев
                IChooseManager mng = doc71.ChooseManager;

                if (mng != null)
                {

                    return mng;
                }
            }
            return null;
        }
        //Получает все объекты из документа
        public IDrawingContainer GetDrawingContainer()
        {
            if (doc7 != null)
            {
                // Получить менеджер видов и слоев
                IViewsAndLayersManager mng = doc7.ViewsAndLayersManager;

                if (mng != null)
                {
                    // Получить коллекцию видов
                    IViews viewsCol = mng.Views;


                    if (viewsCol != null)
                    {
                        // Получить активный вид
                        IView view = viewsCol.ActiveView;

                        if (view != null)
                        {
                            // Получить контейнер графических объектов
                            IDrawingContainer drawCont = (IDrawingContainer)view;
                            return drawCont;
                        }
                    }
                }
            }
            return null;
        }

        //-------------------------------------------------------------------------------
        //  Получить контейнер обозначений 2D
        // ---
        public ISymbols2DContainer GetSymbols2DContainer()
        {
            if (doc7 != null)
            {
                // Получим менеджер для работы с видами и слоями
                ViewsAndLayersManager viewsMng = doc7.ViewsAndLayersManager;

                if (viewsMng != null)
                {
                    // Получим коллекцию видов
                    IViews views = viewsMng.Views;

                    if (views != null)
                    {
                        // Получаем контейнер у активного вида
                        IView view = views.ActiveView;

                        if (view != null)
                            return (ISymbols2DContainer)view;
                    }
                }
            }
            return null;
        }


        public void CloseDocOver(int value)
        {
            while (documents7.Count > value)
            {
                IKompasDocument2D tempDoc = (IKompasDocument2D)documents7[0];
                tempDoc.Close(DocumentCloseOptions.kdSaveChanges);
            }
        }

        //Меняет масштаб
        public bool Mashtab(bool Cheked)
        {
            ILayer layer = KmpsAppl.Doc.Macro.HideLayer(99, true);
            ISelectionManager selection = GetSelectContainer();
            if (selection.SelectedObjects != null)
            {

                SelectConstraintDell(selection);

                doc5.ksLayer(0);

                    if ((selection != null) && (selection.SelectedObjects != null))
                    {
                        double koefX = 0, koefY = 0;

                        if (Cheked == false)
                        {
                            koefX = (100 - KVariable.Give("koefX", "")) / 100;
                            koefY = (100 - KVariable.Give("koefY", "")) / 100;
                        }
                        else
                        {
                            koefX = 100 / (100 - KVariable.Give("koefX", ""));
                            koefY = 100 / (100 - KVariable.Give("koefY", ""));
                        }


                        IDrawingGroup TempGroup = (IDrawingGroup)KmpsAppl.KompasAPI.TransferReference(doc5.ksNewGroup(0), doc5.reference);

                        // Получить массив объектов
                        try
                        {
                            Array arrS = (Array)selection.SelectedObjects;
                            // Если массив есть
                            TempGroup.AddObjects(arrS);

                        }
                        catch
                        {
                            //если один объект
                            object pObj = selection.SelectedObjects;
                            TempGroup.AddObjects(pObj);
                        }

                        IDrawingContainer drawingContainer = GetDrawingContainer();

                        //Добавляем макрообъекты
                        foreach (object obj in drawingContainer.MacroObjects)
                            TempGroup.AddObjects(obj);

                        doc5.ksMtr(0, 0, 0, koefX, koefY);

                        doc5.ksTransformObj(TempGroup.Reference);
                        TempGroup.DetachObjects(TempGroup.Objects[0], true);

                        //Удаляем область трансформирования
                        doc5.ksDeleteMtr();
                        KVariable.UpdateAsync("Angle", 0, "");

                        macro.HideLayer(99, true);
                    }                 
            }
            return !Cheked;

        }

        public void ChangeSelectDimention(IDrawingObject DimObj)
        {
            if (SelectDimenetion != null)
                SelectDimenetion(this, DimObj);
        }

        public ILayer GiveLayer(int number)
        {
            ViewsAndLayersManager ViewsMng = KmpsAppl.Doc.D7.ViewsAndLayersManager;
            IViews views = ViewsMng.Views;
            IView view = views.ActiveView;

            ILayers Layers = view.Layers;
            ILayer layer = Layers.LayerByNumber[number];
            if (layer == null)
            {
                layer = Layers.Add();
                layer.LayerNumber = number;
                layer.Update();
            }

            return layer;
        }

        public dynamic GiveSelectOrChooseObj(bool select = true)
        {
            if (GetSelectContainer().SelectedObjects != null && select)
                return GetSelectContainer().SelectedObjects;
            else if (GetChooseContainer().ChoosenObjects != null)
                    return GetChooseContainer().ChoosenObjects;
            return null;
        }

        public void VisibleLayer(int number, bool visible)
        {
            ViewsAndLayersManager ViewsMng = KmpsAppl.Doc.D7.ViewsAndLayersManager;
            IViews views = ViewsMng.Views;
            IView view = views.ActiveView;

            ILayers Layers = view.Layers;
            ILayer layer = Layers.LayerByNumber[number];
            if (layer != null)
            {
                layer.Visible = visible;
                layer.Update();
            }
        }

        public async static void LockedLayerAsync(int number, bool locked, bool inverse = false)
        {
            await Task.Run(() => { 
            ViewsAndLayersManager ViewsMng = KmpsAppl.Doc.D7.ViewsAndLayersManager;
            IViews views = ViewsMng.Views;
            IView view = views.ActiveView;

            ILayers Layers = view.Layers;
            ILayer layer = Layers.LayerByNumber[number];
            if (layer != null)
            {
                if (inverse)
                    layer.Background = !layer.Background;
                else
                    layer.Background = locked;                    
                layer.Update();
            }
            });
        }

        public void SelectConstraintDell(ISelectionManager selectionManager)
        {
            ISelectionManager selection = GetSelectContainer();
            if (selection.SelectedObjects != null)
            {
                // Получить массив объектов
                try
                {
                    Array arrS = (Array)selection.SelectedObjects;
                    // Если массив есть
                    foreach (object obj in arrS)
                        DellConstraint(obj);

                }
                catch
                {
                    //если один объект
                    object pObj = selection.SelectedObjects;
                    DellConstraint(pObj);
                }

                void DellConstraint (object obj)
                {
                    IDrawingObject1 pDrawObj1 = (IDrawingObject1)obj;
                    pDrawObj1.DeleteConstraints();
                }
            }
        }

        //создаем имя для файла
        public string GiveMeText(string text, ItemCollection IC, string suffix, string Index)
        {
            bool index = (Index != string.Empty);

                Dictionary<string, string> keys = new Dictionary<string, string> 
                {
                    { "{Date}", DateTime.Now.ToString("dd.MM") },
                    { "{Time}", DateTime.Now.ToString("HH:mm") },
                    { "{Number}", KVariable.Give("Number", string.Empty).ToString() },
                    { "{Suffix}", suffix },

                    { "{Sqare}", (index ? KVariable.Give("Sqare", Index) : KVariable.Sum("Sqare", IC)).ToString() },
                    { "{SqareU}", (index ? KVariable.Give("SqareU", Index) : KVariable.Sum("SqareU", IC)).ToString() },
                    { "{Perimetr}", (index ? KVariable.Give("Perimetr", Index) : KVariable.Sum("Perimetr", IC)).ToString() },
                    { "{PerimetrU}", (index ? KVariable.Give("PerimetrU", Index) : KVariable.Sum("PerimetrU", IC)).ToString() },
                    { "{LineP}", (index ? KVariable.Give("LineP", Index) : KVariable.Sum("LineP", IC)).ToString() },
                    { "{CurveP}", (index ? KVariable.Give("CurveP", Index) : KVariable.Sum("CurveP", IC)).ToString() },
                    { "{Shov}", (index ? KVariable.Give("Shov", Index) : KVariable.Sum("Shov", IC)).ToString() },
                    { "{photo}", KVariable.Give("photo", string.Empty).ToString() },
                    { "{photoNumber}", KVariable.GiveNote("photo", string.Empty) },

                    { "{lenth}", (index ? KVariable.Give("lenth", Index) : KVariable.Sum("lenth", IC)).ToString() },
                    { "{cut}", (index ? KVariable.Give("cut", Index) : KVariable.Sum("cut", IC)).ToString() },

                    { "{Angle}", KVariable.Give("Angle", string.Empty).ToString() },

                    { "{koefX}", KVariable.Sum("koefX", IC).ToString() },
                    { "{koefY}", KVariable.Sum("koefY", IC).ToString() },

                    { "{Comment1}", KVariable.GiveNote("Comment1", string.Empty) },
                    { "{Comment2}", KVariable.GiveNote("Comment2", string.Empty) },
                    { "{Comment3}", KVariable.GiveNote("Comment3", string.Empty) },

                };
            if (index)
            {
                keys.Add("{color}", KVariable.GiveNote("color", Index));
                keys.Add("{factura}", KVariable.GiveNote("factura", Index));
            }

            return changeTagWithText(text, keys);
        }


        public static string changeTagWithText(string text, Dictionary<string, string> dict)
        {
            foreach (KeyValuePair<string, string> entry in dict)
                text = text.Replace(entry.Key, entry.Value);
            return text;
        }

        public void CreateQrCode(bool name, double scale, double x = 0, double y = 0)
        {
            if (KmpsAppl.Appl.ActiveDocument.Path != string.Empty)
            {
                ksRequestInfo info = (ksRequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                if ((x == 0) && (y == 0))
                    KmpsAppl.Doc.D5.ksCursor(info, ref x, ref y, 0);

                Bitmap bitmap = OnEncode(name ? KmpsAppl.Appl.ActiveDocument.Name : KmpsAppl.Appl.ActiveDocument.PathName);
                string outputFileName = Path.GetTempFileName().Replace(".tmp", ".jpg");

                using (MemoryStream memory = new MemoryStream())
                {
                    using (FileStream fs = new FileStream(outputFileName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        bitmap.Save(memory, ImageFormat.Jpeg);
                        byte[] bytes = memory.ToArray();
                        fs.Write(bytes, 0, bytes.Length);
                    }
                }
                IRasters rasters = KmpsAppl.Doc.GetDrawingContainer().Rasters;
                IRaster raster = rasters.Add();
                raster.FileName = outputFileName;
                raster.SetPlacement(x, y - scale, 0, false);
                raster.Scale = scale / (bitmap.Height/3.77);
                raster.InsertionType = true;
                raster.Update();

                Bitmap OnEncode(string Data)
                {
                    QRCodeGenerator qrGenerator = new QRCodeGenerator();
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(Data,
                    QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeImage = qrCode.GetGraphic(20);
                    return qrCodeImage;
                }
            }
        }

        public async void CreateText(string text, double x, double y, double width, ItemCollection IC, string suffix, bool autoSize, bool stroke, bool table, double textSize = 10)
        {

            IDrawingTexts drawingTexts = GetDrawingContainer().DrawingTexts;
            IDrawingText drawingText = drawingTexts.Add();
            IText txt = (IText)drawingText;
            drawingText.X = x;
            drawingText.Y = y;
            drawingText.Allocation = ksAllocationEnum.ksAlLeft;

            Regex BeginMatch = new Regex("<([^/>].*?)>"); //Открывающие тэги
            Regex EndMatch = new Regex("</(.*?)>"); //Закрывающие

            string[] strSplit = text.Split('\n');
            KmpsAppl.ProgressBar.Start(0, strSplit.Length, "Тэги", true);
            for (int s = 0; s < strSplit.Length; s++)
            {
                KmpsAppl.ProgressBar.SetProgress(s,"Строка: " + s, false);
                ITextLine textLine = txt.Add();
                MakeTextFromTag(strSplit[s], string.Empty, textLine);
            }
            KmpsAppl.ProgressBar.Stop("Закончили", true);


            //Насильно выставляем размер шрифта. Это пиздец, но только так работает.
            KmpsAppl.ProgressBar.Start(0, txt.Count, "Размер", true);
            foreach (ITextLine line in txt.TextLines)
            {
                await Task.Run(() =>
                {
                    List<ITextItem> ItemList = new List<ITextItem>();
                    foreach (ITextItem item in line.TextItems)
                    {
                        ITextFont textFont = (ITextFont)item;
                        textFont.Height = autoSize ? width / 30 : textSize;
                        ItemList.Add(item);
                    }

                    for (int Li = 0; Li < ItemList.Count; Li++)
                    {
                        ItemList[Li].Update();
                        drawingText.Update();
                    }
                });
            }
            KmpsAppl.ProgressBar.Stop("Закончили", true);

            if (KVariable.Give("photo", string.Empty) > 0 && table)
            {
                double tW = width * 1; //540
                double tH = tW * 0.15; //100
                double szText = autoSize ? width / 30 : textSize;

                y += 50;
                doc5.ksTable();
                //Горизонтальные линии
                doc5.ksLineSeg(x, y + tH, x + tW, y + tH, 1); //150
                doc5.ksLineSeg(x, y + 0.6 * tH, x + tW, y + 0.6 * tH, 1); //110
                doc5.ksLineSeg(x, y + 0.4 * tH, x + tW, y + 0.4 * tH, 2);  //90
                doc5.ksLineSeg(x, y + 0.2 * tH, x + tW, y + 0.2 * tH, 2);  //70
                doc5.ksLineSeg(x, y, x + tW, y, 1); //0
                                                    //Вертикальные
                doc5.ksLineSeg(x, y + tH, x, y, 1);
                doc5.ksLineSeg(x + 0.16 * tW, y + tH, x + 0.16 * tW, y, 2);
                doc5.ksLineSeg(x + 0.33 * tW, y + tH, x + 0.33 * tW, y, 2);
                doc5.ksLineSeg(x + 0.54 * tW, y + tH, x + 0.54 * tW, y, 2);
                doc5.ksLineSeg(x + 0.66 * tW, y + tH, x + 0.66 * tW, y, 2);
                doc5.ksLineSeg(x + 0.83 * tW, y + tH, x + 0.83 * tW, y, 2);
                doc5.ksLineSeg(x + tW, y + tH, x + tW, y, 1);

                //Шапка
                doc5.ksText(x + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "№ Заказа");
                doc5.ksText(x + 0.16 * tW + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "№ Счета");
                doc5.ksText(x + 0.33 * tW + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "Материал");
                doc5.ksText(x + 0.54 * tW + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "Ролик");
                doc5.ksText(x + 0.66 * tW + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "Остаток");
                doc5.ksText(x + 0.83 * tW + 2, y + 0.6 * tH + 2, 0, szText, 1, 0, "Причина");
                //Вторая строка
                doc5.ksText(x + 2, y + 0.4 * tH + 2, 0, szText, 1, 0, KVariable.Give("Number", string.Empty).ToString());
                doc5.ksText(x + 0.16 * tW + 2, y + 0.4 * tH + 2, 0, szText, 1, 0, KVariable.GiveNote("Comment3", string.Empty));
                doc5.ksText(x + 0.33 * tW + 2, y + 0.4 * tH + 2, 0, szText, 1, 0, KVariable.GiveNote("factura", IC[0].ToString()) + " " + KVariable.GiveNote("color", IC[0].ToString()));
                doc5.ksEndObj();
            }


            async void MakeTextFromTag(string str, string index, ITextLine line, bool bold = false, bool italic = true, bool loop = false)
            {
                int temp = txt.Count;

                List<string> splitStroke = SplitTag(str);

                for (int i = splitStroke.Count - 1; i >= 0; i--)
                {
                        Match match = BeginMatch.Match(splitStroke[i]);
                    if (match.Length > 0)
                    {
                        try
                        {
                            switch (match.Groups[1].Value)
                            {
                                case "b":
                                    MakeTextFromTag(splitStroke[i].Substring(match.Length), index, line, true, italic, loop);
                                    break;
                                case "i":
                                    break;
                                case "loop":
                                    for (int ic = 0; ic < IC.Count; ic++)
                                        MakeTextFromTag(splitStroke[i].Substring(match.Length) + "\n", IC[ic].ToString(), line, bold, italic, loop);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("if:" + ex.Message + "\n i=" + i + " Count=" + splitStroke.Count);
                        }
                    }
                    else
                    {
                        try
                        {

                            ITextItem textItem = line.AddBefore(0);

                        ITextFont fontItem = (ITextFont)textItem;
                        fontItem.Bold = bold;
                        fontItem.Italic = italic;

                        textItem.Str = GiveMeText(splitStroke[i].Replace("\r", string.Empty), IC, suffix, index);
                        textItem.Update();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("else:" + ex.Message + "\n i=" + i + " Count=" + splitStroke.Count);
                        }

                    }
                }

                List<string> SplitTag(string SplitStr, int lastPosition = 0)
                {
                    MatchCollection matchesBegin = BeginMatch.Matches(str);
                    MatchCollection matchesEnd = EndMatch.Matches(str);

                    List<string> SplitStrs = new List<string>();
                    if (matchesBegin.Count > 0)
                    {
                        for (int m = 0; m < matchesBegin.Count; m++)
                        {
                            
                            //Если в начале текста есть что то без тэга
                            if (matchesBegin[m].Index - lastPosition > 0)
                                SplitStrs.Add(SplitStr.Substring(lastPosition, matchesBegin[m].Index - lastPosition));
                            //Добавляем строку внутри тэга

                            //для наглядности. Получаем индекс закрывающего тега

                                int closeIndex = returnCloseTagIndex(matchesBegin[m].Groups[1].Value, m, matchesEnd);
                            SplitStrs.Add(SplitStr.Substring(matchesBegin[m].Index, matchesEnd[closeIndex].Index - matchesBegin[m].Index));

                            lastPosition = matchesEnd[closeIndex].Index + matchesEnd[closeIndex].Length;
                                m = closeIndex;
                            //Если вконце есть текст без тэгов
                            if (m == matchesEnd.Count - 1)
                                if (SplitStr.Length - lastPosition > 0) SplitStrs.Add(SplitStr.Substring(lastPosition, SplitStr.Length - lastPosition));

                        }
                    }
                    else SplitStrs.Add(SplitStr);
                    return SplitStrs;

                    int returnCloseTagIndex(string value, int openTagIndex, MatchCollection match)
                    {
                        for (int mc = openTagIndex; mc < match.Count; mc++)
                            if (match[mc].Groups[1].Value == value)
                                return mc;
                        return -1;
                    }
                }
            }
        }

        public void DeleteSelectObj()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                dynamic objects = GetSelectContainer().SelectedObjects;
                if (objects != null)
                {
                    try
                    {
                        foreach (IDrawingObject drawingObject in objects)
                            drawingObject.Delete();
                    }
                    catch
                    {
                        ((IDrawingObject)objects).Delete();
                    }
                }
            }
            
        }

       
    }
}
