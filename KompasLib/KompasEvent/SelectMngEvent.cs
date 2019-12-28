using Kompas6API5;
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
