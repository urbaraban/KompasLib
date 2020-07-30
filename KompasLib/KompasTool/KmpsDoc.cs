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
        private KmpsAppl kmpsAppl;

        public KVariable Var;

        public SizeTool ST { get; }

        public KmpsAttr Attribute { get; }


        public IKompasDocument2D D7 { get; set; }

        public IKompasDocument1 D1 { get; }


        public IKompasDocument2D1 D71 { get; }

        public ksDocument2D D5 { get; set; }

        public IDocuments Documents7 { get; set; }

        public ksDocumentParam DocPar { get; set; }

        public KmpsMacro Macro { get; set; }


        public KmpsDoc(KmpsAppl kmps)
        {
            this.kmpsAppl = kmps;
            if (KmpsAppl.KompasAPI != null)
            {
                this.Documents7 = KmpsAppl.Appl.Documents;

                // Получаем интерфейс активного документа 2D в API7
                this.D7 = (IKompasDocument2D)KmpsAppl.Appl.ActiveDocument;

                if (D7 == null)
                    return;

                this.D71 = (IKompasDocument2D1)KmpsAppl.Appl.ActiveDocument;
                this.D1 = (IKompasDocument1)KmpsAppl.Appl.ActiveDocument;

                Attribute = new KmpsAttr(this.kmpsAppl);

                Attribute.FuncAttrType();

                // Получаем интерфейс активного документа 2D в API5
                D5 = (ksDocument2D)KmpsAppl.KompasAPI.ActiveDocument2D();
                if (D5 == null)
                    return;

                this.DocPar = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
                if (DocPar == null)
                    return;
                else this.DocPar.type = (short)DocType.lt_DocFragment;

                this.Macro = new KmpsMacro(this.kmpsAppl);
                this.ST = new SizeTool(this);

                this.Var = new KVariable(this);
            }
        }


        //Получает только селектированные объекты
        public ISelectionManager GetSelectContainer()
        {
            if (D71 != null)
            {
                // Получить менеджер видов и слоев
                ISelectionManager mng = D71.SelectionManager;

                if (mng != null)
                    return mng;
            }
            return null;
        }

        //Получает только выделенные объекты
        public IChooseManager GetChooseContainer()
        {
            if (D71 != null)
            {
                // Получить менеджер видов и слоев
                IChooseManager mng = D71.ChooseManager;

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
            if (D7 != null)
            {
                // Получить менеджер видов и слоев
                IViewsAndLayersManager mng = D7.ViewsAndLayersManager;

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
            if (D7 != null)
            {
                // Получим менеджер для работы с видами и слоями
                ViewsAndLayersManager viewsMng = D7.ViewsAndLayersManager;

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


        public IDrawingGroup GetDrawingGroup()
        {
            return  (IDrawingGroup)KmpsAppl.KompasAPI.TransferReference(this.D5.ksNewGroup(0), this.D5.reference);
        }

        public ILayer GiveLayer(int number)
        {
            ViewsAndLayersManager ViewsMng = this.kmpsAppl.Doc.D7.ViewsAndLayersManager;
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
            ViewsAndLayersManager ViewsMng = this.kmpsAppl.Doc.D7.ViewsAndLayersManager;
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

        public async void LockedLayerAsync(int number, bool locked, bool inverse = false)
        {
            await Task.Run(() => { 
            ViewsAndLayersManager ViewsMng = this.kmpsAppl.Doc.D7.ViewsAndLayersManager;
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

        public void CreateQrCode(string message, double scale, double x = 0, double y = 0)
        {
            if (KmpsAppl.Appl.ActiveDocument.Path != string.Empty)
            {
                ksRequestInfo info = (ksRequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                if ((x == 0) && (y == 0))
                    this.kmpsAppl.Doc.D5.ksCursor(info, ref x, ref y, 0);

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
                IRasters rasters = this.kmpsAppl.Doc.GetDrawingContainer().Rasters;
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
