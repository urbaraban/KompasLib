using Kompas6API5;
using Kompas6API7;
using reference = System.Int32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KAPITypes;

namespace KompasLib.Tools
{
    public class KmpsMacro
    {
        KmpsAppl api;

        public KmpsMacro(KmpsAppl API)
        {
            api = API;
        }

        //
        //Обращение к макрообъектам
        //
        //Создание
        public IMacroObject MakeCeilingMacro(string index)
        {
            api.Doc.D5.ksLayer(99);
            IMacroObject pMacroObj;
            ILayer layer = HideLayer(99, true);
            api.Doc.D5.ksMacro(1);
            reference macroRef = api.Doc.D5.ksEndObj();
            pMacroObj = (IMacroObject)KmpsAppl.KompasAPI.TransferReference(macroRef, api.Doc.D5.reference);
            pMacroObj.Name = "Ceiling:" + index;
            pMacroObj.Update();
            api.Doc.D5.ksLayer(0);
            return pMacroObj;
        }

        //хуй знает зачем я теперь получаю слой, но эта функция его скрывает.
        public ILayer HideLayer(int number, bool status)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IViewsAndLayersManager ViewsAndLayersManager = api.Doc.D7.ViewsAndLayersManager;
                IViews Views = ViewsAndLayersManager.Views;
                //получаем текущий вид
                IView CurView = Views.ActiveView;
                ILayers Layers = CurView.Layers;
                ILayer layer = Layers.LayerByNumber[number];
                if (layer != null)
                {
                    layer.Visible = !status;
                    layer.Update();
                }
                return layer;
            }
            return null;
        }

        //Добавление в макрообъект
        public bool AddCeilingMacro(reference refObj, string index)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                object pDrawObj = (object)KmpsAppl.KompasAPI.TransferReference(refObj, api.Doc.D5.reference);
                if (pDrawObj != null)
                {
                    IMacroObject pMacroObj = FindCeilingMacro(index);
                    pMacroObj.AddObjects(pDrawObj);
                    pMacroObj.Update();
                    return true;
                }
            }
            return false;
        }

        //Поиск макрообъекта по имени
        public IMacroObject FindCeilingMacro(string index)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IMacroObject pMacroObj = null;
                IDrawingContainer drawingContainer = api.Doc.GetDrawingContainer();
                try
                {
                    MacroObjects array = (MacroObjects)drawingContainer.MacroObjects;

                    foreach (IMacroObject obj in array)
                        if (obj.Name == "Ceiling:" + index) return obj;
                }
                catch
                {
                    MacroObject pObj = (MacroObject)drawingContainer.MacroObjects;
                    pMacroObj = (IMacroObject)pObj;
                }
                return pMacroObj;
            }
            return null;
        }

        //Удаление макрообъекта
        public bool RemoveCeilingMacro(string index)
        {
            if (api.Doc.D5 != null)
            {
                IMacroObject pMacroObj = FindCeilingMacro(index);
                if (pMacroObj != null) return pMacroObj.Delete();
                else return false;
            }
            return false;
        }

        //Получаем составляющий контур из ksMakeEncloseContour для рассчета периметра
        public IContour GiveContour(reference refContour)
        {
            IContour contour = null;
            IDrawingObject pDrawObj = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(refContour, api.Doc.D5.reference);
            if (pDrawObj != null)
            {
                if ((int)pDrawObj.DrawingObjectType == 26)
                {
                    contour = (IContour)pDrawObj;
                }
            }
            //ksIterator.ksDeleteIterator();
            return contour;
        }

        //Получаем референс ksMakeEncloseContour из Макрообъекта для рассчета площади.
        public reference GiveRefFromMacro(reference refMacro, string position)
        {
            ksIterator ksIterator = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
            ksIterator.ksCreateIterator(ldefin2d.ALL_OBJ, refMacro);
            reference refObj = ksIterator.ksMoveIterator(position);
            ksIterator.ksDeleteIterator();
            return refObj;
        }

        public void ClearRefFromMacro(reference refMacro)
        {
            ksIterator ksIterator = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
            ksIterator.ksCreateIterator(ldefin2d.ALL_OBJ, refMacro);
            reference refObj = ksIterator.ksMoveIterator("F");
            api.Doc.D5.ksDeleteObj(refObj);
            ksIterator.ksDeleteIterator();
        }

    }
}
