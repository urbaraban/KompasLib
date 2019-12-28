////////////////////////////////////////////////////////////////////////////////
//
// DocumentEvent  - обработчик событий от документа
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__

#if (__LIGHT_VERSION__)
	using Kompas6LTAPI5;
	using KompasLTAPI7;
#else
using Kompas6API5;
using Kompas6API7;
//	using KompasAPI7;
#endif
using KompasLib.Tools;


namespace KompasLib.Event
{
    public class DocumentEvent : BaseEvent, Kompas6API5.ksDocumentFileNotify
    {
        private KmpsAppl api;
        public DocumentEvent(object doc, KmpsAppl API)
            : base(doc, typeof(Kompas6API5.ksDocumentFileNotify).GUID, doc,
            -1)
        { this.api = API; }


        // kdBeginCloseDocument - Начало закрытия документа
        public bool BeginCloseDocument()
        {
            return true;
        }


        // kdCloseDocument - Документ закрыт
        public bool CloseDocument()
        {
            TerminateEvents(null, m_Doc, -1, null);
            return true;
        }


        // kdBeginSaveDocument - Начало сохранения документа
        public bool BeginSaveDocument(string fileName)
        {
            return true;
        }


        // kdSaveDocument - Документ сохранен
        public bool SaveDocument()
        {
            return true;
        }


        // kdActiveDocument - Документ активизирован.
        public bool Activate()
        {

            return true;
        }


        // kdDeactiveDocument - Документ деактивизирован.
        public bool Deactivate()
        {
            return true;
        }


        // kdBeginSaveAsDocument - Начало сохранения документа c другим именем (до диалога выбора имени)
        public bool BeginSaveAsDocument()
        {
            return true;
        }


        // kdDocumentFrameOpen - Окно документа открылось
        public bool DocumentFrameOpen(object v)
        {
            if (!FindEvent(typeof(DocumentFrameEvent), (object)m_Container, 0))
            {
                if (v != null)
                {
                    DocumentFrameEvent frmEvent = new DocumentFrameEvent((ksDocumentFrameNotify)v, (object)m_Container, true);
                    frmEvent.Advise();
                }
            }
            return true;
        }

        public bool ProcessActivate(int id)
        {
            return true;
        }

        public bool ProcessDeactivate(int id)
        {
            return true;
        }

        public bool BeginProcess(int iD)
        {
            return true;
        }
        public bool EndProcess(int iD, bool Success)
        {
            return true;
        }
    }
}
