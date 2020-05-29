using Kompas6API5;
using Kompas6API7;
using KompasLib.Event;
using KompasLib.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KompasLib.Event
{
    public class SelectMngEvent : BaseEvent, ksSelectionMngNotify
    {
        public SelectMngEvent(object obj, object doc, bool selfAdvise)
            : base(obj, typeof(ksSelectionMngNotify).GUID, doc,
            -1)
        { }

        // ksmSelect - Объект селектирован
        public bool Select(object obj)
        {
            try
            {
                Array array = KmpsAppl.Doc.GetSelectContainer().SelectedObjects;
            }
            catch
            {
                IDrawingObject drawingObject = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference((int)obj, KmpsAppl.Doc.D5.reference);
                switch (drawingObject.DrawingObjectType)
                {
                    case Kompas6Constants.DrawingObjectTypeEnum.ksDrLDimension:
                        KmpsAppl.Doc.ChangeSelectDimention(drawingObject);
                        break;
                    case Kompas6Constants.DrawingObjectTypeEnum.ksDrADimension:
                        KmpsAppl.Doc.ChangeSelectDimention(drawingObject);
                        break;
                }
            }
            return true;
        }


        // ksmUnselect - Объект расселектирован
        public bool Unselect(object obj)
        {
            return true;
        }


        // ksmUnselectAll - Все объекты расселектированы
        public bool UnselectAll()
        {
            return true;
        }
    }
}
