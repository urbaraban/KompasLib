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

namespace KompasLib.Tools
{
    public class KmpsDoc
    {
        private ksDocument2D doc5;
        private IKompasDocument2D doc7;
        private IKompasDocument2D1 doc71;
        private IKompasDocument1 doc1;
        private ksDocumentParam docPar;
        private IDocuments documents7;
        private SizeTool sizeTool;
        private KmpsMacro macro;
        private KmpsAttr attr;




        public KVariable Var;

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

                this.doc71 = (IKompasDocument2D1)KmpsAppl.Appl.ActiveDocument;
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

                this.Var = new KVariable(this);

            }
        }


        public IKompasDocument2D D7
        {
            get => this.doc7;
            set => this.doc7 = value;
        }

        public IKompasDocument1 D1
        {
            get => this.doc1;
            set => this.doc1 = value;
        }


        public IKompasDocument2D1 D71
        {
            get => this.doc71;
        }

        public ksDocument2D D5
        {
            get => this.doc5;
            set => this.doc5 = value;
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
                        koefX = (100 - this.Var.Variable("koefX", "", false).Value) / 100;
                        koefY = (100 - this.Var.Variable("koefY", "").Value) / 100;
                    }
                    else
                    {
                        koefX = 100 / (100 - this.Var.Variable("koefX", "").Value);
                        koefY = 100 / (100 - this.Var.Variable("koefY", "").Value);
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
 




        public void CreateQrCode(string message, double scale, double x = 0, double y = 0)
        {
            if (KmpsAppl.Appl.ActiveDocument.Path != string.Empty)
            {
                ksRequestInfo info = (ksRequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                if ((x == 0) && (y == 0))
                    KmpsAppl.Doc.D5.ksCursor(info, ref x, ref y, 0);

                Bitmap bitmap = OnEncode(message);
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
