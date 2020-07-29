////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - обработчик событий от приложения
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__

using Kompas6API5;
using Kompas6API7;
using KompasLib.Tools;
using System;
using System.Threading;

namespace KompasLib.Event
{
    public class ApplicationEvent : BaseEvent, ksKompasObjectNotify
    {
        public event EventHandler<object> OpenedDoc;
        public event EventHandler<object> CreatedDoc;
        public event EventHandler<object> ChangeDoc;
        public event EventHandler<Tuple<int, bool>> EvKeyUp;
        public event EventHandler DisconnectDoc;

        public ApplicationEvent(object obj)
          : base(obj, typeof(ksKompasObjectNotify).GUID,
          null, -1)
        { }


        // koApplicatinDestroy - Закрытие приложения
        public bool ApplicationDestroy()
        {
            // Самоудаление
            TerminateEvents();
            DisconnectDoc?.Invoke(this, null);
            return true;
        }


        // koBeginCloseAllDocument - Начало закрытия всех открытых документов
        public bool BeginCloseAllDocument()
        {
            bool res = true;
            return res;
        }


        // koBeginCreate - Начало создания документа(до диалога выбора типа)
        public bool BeginCreate(int docType)
        {
            return true;
        }


        // koOpenDocumenBegin - Начало открытия документа
        public bool BeginOpenDocument(string fileName)
        {
            return true;
        }


        // koBeginOpenFile - Начало открытия документа(до диалога выбора имени)
        public bool BeginOpenFile()
        {
            return true;
        }


        // koActiveDocument - Переключение на другой активный документ
        public bool ChangeActiveDocument(object newDoc, int docType)
        {
            ChangeDoc?.Invoke(this, newDoc);
            return true;
        }


        // koCreateDocument - Документ создан
        public bool CreateDocument(object newDoc, int docType)
        {
            CreatedDoc?.Invoke(this, newDoc);
            return true;
        }


        // koOpenDocumen - Документ открыт
        public bool OpenDocument(object newDoc, int docType)
        {
            OpenedDoc?.Invoke(null, newDoc);
            return true;
        }

        // koKeyDown - Событие клавиатуры
        public bool KeyDown(ref int key, int flags, bool system)
        {
            return true;
        }

        // koKeyUp - Событие клавиатуры
        public bool KeyUp(ref int key, int flags, bool system)
        {
            EvKeyUp?.Invoke(this, new Tuple<int, bool>(key, system));
            return true;
        }

        // koKeyPress - Событие клавиатуры
        public bool KeyPress(ref int key, bool system)
        {
            return true;
        }

        public bool BeginReguestFiles(int type, ref object files)
        {
            return true;
        }

        public bool BeginChoiceMaterial(int MaterialPropertyId)
        {
            return true;
        }
        public bool ChoiceMaterial(int MaterialPropertyId, string material, double density)
        {
            return true;
        }
    }
}
