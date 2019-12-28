using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace KompasLib.Tools
{
    public class KmpsDoc
    {
        private KmpsAppl api;
        private ksDocument2D doc5;
        private IKompasDocument2D doc7;
        private IKompasDocument2D1 doc71;
        private IKompasDocument1 doc1;
        private ksDocumentParam docPar;
        private IDocuments documents7;
        private KmpsVariable variable;
        private SizeTool sizeTool;
        private KmpsMacro macro;

        KmpsAttr attr;




        public SizeTool ST
        {
            get => sizeTool;
        }

        public KmpsAttr Attribute
        {
            get => attr;
        }

        public KmpsDoc(KmpsAppl API)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                this.api = API;

                this.documents7 = KmpsAppl.Appl.Documents;

                // Получаем интерфейс активного документа 2D в API7
                this.doc7 = (IKompasDocument2D)KmpsAppl.Appl.ActiveDocument;

                if (doc7 == null)
                    return;



                this.doc71 = (IKompasDocument2D1)KmpsAppl.Appl.ActiveDocument;
                this.doc1 = (IKompasDocument1)KmpsAppl.Appl.ActiveDocument;

                attr = new KmpsAttr(API);
                
                attr.FuncAttrType();

                // Получаем интерфейс активного документа 2D в API5
                this.doc5 = (ksDocument2D)KmpsAppl.KompasAPI.ActiveDocument2D();
                if (doc5 == null)
                    return;

                this.docPar = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
                if (docPar == null)
                    return;
                else this.docPar.type = (short)DocType.lt_DocFragment;



                this.variable = new KmpsVariable(doc71);
                this.macro = new KmpsMacro(api);
                this.sizeTool = new SizeTool(this, api);
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


        public IKompasDocument2D1 D71
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

        public KmpsVariable Variable 
        {
            get => variable;
            set => variable = value;
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
                {

                    return mng;
                }
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

        public bool Mashtab(bool Cheked)
        {

            ILayer layer = api.Doc.Macro.HideLayer(99, true);
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
                        koefX = (100 - variable.Give("koefX", "")) / 100;
                        koefY = (100 - variable.Give("koefY", "")) / 100;
                    }
                    else
                    {
                        koefX = 100 / (100 - variable.Give("koefX", ""));
                        koefY = 100 / (100 - variable.Give("koefY", ""));
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
                    variable.Update("Angle", 0, "");

                    macro.HideLayer(99, true);

                    
                }
            }
            return !Cheked;
        }

        public ILayer GiveLayer(int number)
        {
            ViewsAndLayersManager ViewsMng = api.Doc.D7.ViewsAndLayersManager;
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

        public void VisibleLayer(int number, bool visible)
        {
            ViewsAndLayersManager ViewsMng = api.Doc.D7.ViewsAndLayersManager;
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
                    { "{Number}", variable.Give("Number", string.Empty).ToString() },
                    { "{Suffix}", suffix },

                    { "{Sqare}", (index ? variable.Give("Sqare", Index) : variable.Sum("Sqare", IC)).ToString() },
                    { "{SqareU}", (index ? variable.Give("SqareU", Index) : variable.Sum("SqareU", IC)).ToString() },
                    { "{Perimetr}", (index ? variable.Give("Perimetr", Index) : variable.Sum("Perimetr", IC)).ToString() },
                    { "{PerimetrU}", (index ? variable.Give("PerimetrU", Index) : variable.Sum("PerimetrU", IC)).ToString() },
                    { "{LineP}", (index ? variable.Give("LineP", Index) : variable.Sum("LineP", IC)).ToString() },
                    { "{CurveP}", (index ? variable.Give("CurveP", Index) : variable.Sum("CurveP", IC)).ToString() },
                    { "{Shov}", (index ? variable.Give("Shov", Index) : variable.Sum("Shov", IC)).ToString() },
                    { "{photo}", (index ? variable.Give("photo", Index) : variable.Sum("photo", IC)).ToString() },
                    { "{lenth}", (index ? variable.Give("lenth", Index) : variable.Sum("lenth", IC)).ToString() },

                    { "{cut}", (index ? variable.Give("cut", Index) : variable.Sum("cut", IC)).ToString() },

                    { "{Angle}", variable.Give("Angle", string.Empty).ToString() },

                    { "{koefX}", variable.Sum("koefX", IC).ToString() },
                    { "{koefY}", variable.Sum("koefY", IC).ToString() },

                    { "{Comment1}", variable.GiveNote("Comment1", string.Empty).ToString() },
                    { "{Comment2}", variable.GiveNote("Comment2", string.Empty).ToString() },

                };
            if (index)
            {
                keys.Add("{color}", variable.GiveNote("color", Index));
                keys.Add("{factura}", variable.GiveNote("factura", Index));
            }

            return changeTagWithText(text, keys);
        }


        public static string changeTagWithText(string text, Dictionary<string, string> dict)
        {
            foreach (KeyValuePair<string, string> entry in dict)
                text = text.Replace(entry.Key, entry.Value);
            return text;
        }

        public void CreateText(string text, double x, double y, double width, ItemCollection IC, string suffix, bool autoSize, bool stroke, double textSize = 10)
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
            api.ProgressBar.Start(0, strSplit.Length, "Тэги", true);
            for (int s = 0; s < strSplit.Length; s++)
            {
                api.ProgressBar.SetProgress(s,"Строка: " + s, false);
                ITextLine textLine = txt.Add();
                MakeTextFromTag(strSplit[s], string.Empty, textLine);
            }
            api.ProgressBar.Stop("Закончили", true);

            //Насильно выставляем размер шрифта. Это пиздец, но только так работает.
            int ti = 0;
            api.ProgressBar.Start(0, txt.Count, "Размер", true);
            foreach (ITextLine line in txt.TextLines)
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
                api.ProgressBar.SetProgress(ti++, "Строка: " + ti, false);
            }
            api.ProgressBar.Stop("Закончили", true);


            void MakeTextFromTag(string str, string index, ITextLine line, bool bold = false, bool italic = true, bool loop = false)
            {

                int temp = txt.Count;

                List<string> splitStroke = SplitTag(str);

                for (int i = splitStroke.Count - 1; i >= 0; i--)
                {
                    Match match = BeginMatch.Match(splitStroke[i]);
                    if (match.Length > 0)
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
                    else
                    {
                        ITextItem textItem = line.AddBefore(0);

                        ITextFont fontItem = (ITextFont)textItem;
                        fontItem.Bold = bold;
                        fontItem.Italic = italic;

                        textItem.Str = GiveMeText(splitStroke[i].Replace("\r", string.Empty), IC, suffix, index);
                        textItem.Update();
                    }
                }

                List<string> SplitTag(string SplitStr, int lastPosition = 0)
                {
                    MatchCollection matchesBegin = BeginMatch.Matches(str);
                    MatchCollection matchesEnd = EndMatch.Matches(str);

                    List<string> SplitStrs = new List<string>();
                    if (matchesBegin.Count > 0)
                    {
                        for (int i = 0; i < matchesBegin.Count; i++)
                        {
                            //Если в начале текста есть что то без тэга
                            if (matchesBegin[i].Index - lastPosition > 0)
                                SplitStrs.Add(SplitStr.Substring(lastPosition, matchesBegin[i].Index - lastPosition));
                            //Добавляем строку внутри тэга

                            //для наглядности. Получаем индекс закрывающего тега

                            int closeIndex = returnCloseTagIndex(matchesBegin[i].Groups[1].Value, i, matchesEnd);
                            SplitStrs.Add(SplitStr.Substring(matchesBegin[i].Index, matchesEnd[closeIndex].Index - matchesBegin[i].Index));

                            lastPosition = matchesEnd[closeIndex].Index + matchesEnd[closeIndex].Length;
                            i = closeIndex;
                            //Если вконце есть текст без тэгов
                            if (i == matchesEnd.Count - 1)
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

            //Бъем текст на строки



        }


    }
}
