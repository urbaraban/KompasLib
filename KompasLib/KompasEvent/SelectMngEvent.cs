using Kompas6API5;
using System;


namespace KompasLib.Event
{
    public class SelectMngEvent : BaseEvent, ksSelectionMngNotify
    {

        public event EventHandler<int> SelectedObject;
        public event EventHandler<int> UnselectedObject;

        public SelectMngEvent(object obj, object doc)
            : base(obj, typeof(ksSelectionMngNotify).GUID, doc,
            -1)
        { }

        // ksmSelect - Объект селектирован
        public bool Select(object obj)
        {
            SelectedObject?.Invoke(this, (int)obj);
            return true;
        }


        // ksmUnselect - Объект расселектирован
        public bool Unselect(object obj)
        {
            UnselectedObject?.Invoke(this, (int)obj);
            return true;
        }


        // ksmUnselectAll - Все объекты расселектированы
        public bool UnselectAll()
        {
            SelectedObject?.Invoke(this, 0);
            return true;
        }
    }
}
