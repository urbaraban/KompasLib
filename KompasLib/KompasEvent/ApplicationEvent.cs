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
        public DocumentOpenedDelegate DocumentOpened { get; set; }
        public delegate void DocumentOpenedDelegate(object newDoc, int docType);

        public DocumentCreateDelegate DocumentCreated { get; set; }
        public delegate void DocumentCreateDelegate(object newDoc, int docType);

        public DocumentChangedDelegate DocumentChanged { get; set; }
        public delegate void DocumentChangedDelegate(object DocDispatch, int docType);

        public KeyUpDelegate KeyboardDown { get; set; }
        public delegate void KeyUpDelegate(int key, int flags, bool system);
       

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

        // koCreateDocument - Документ создан
        public bool CreateDocument(object newDoc, int docType)
        {
            DocumentCreated?.Invoke(newDoc, docType);
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
            DocumentChanged?.Invoke(newDoc, docType);
            return true;
        }


        // koOpenDocumen - Документ открыт
        public bool OpenDocument(object newDoc, int docType)
        {
            DocumentOpened?.Invoke(newDoc, docType);
            return true;
        }

        // koKeyDown - Событие клавиатуры
        public bool KeyDown(ref int key, int flags, bool system)
        {
            KeyboardDown?.Invoke(key, flags, system);
            return true;
        }

        // koKeyUp - Событие клавиатуры
        public bool KeyUp(ref int key, int flags, bool system)
        {
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
