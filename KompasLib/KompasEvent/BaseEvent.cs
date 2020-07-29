/////////////////////////////////////////////////////////////////////////////
//
// Базовый клас для обработчиков событий
//
/////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__

using Kompas6API5;

using System;
using System.Resources;
using System.Diagnostics;
using System.Collections;
using System.Runtime.InteropServices.ComTypes;
using Kompas6Constants;
using KAPITypes;
using KompasLib.Tools;

namespace KompasLib.Event
{
    public class BaseEvent
    {
        protected int m_Cookie;
        protected object m_Container;
        protected Guid m_Events;
        protected object m_Doc;
        protected int m_ObjType;
        protected IConnectionPoint m_ConnPt;
        protected string m_LibName;
        protected string str = string.Empty;

        public BaseEvent(object obj, Guid events, object doc, int objType
            )
        {
            m_Cookie = 0;
            m_Container = obj;
            m_Events = events;
            m_Doc = doc;
            m_ObjType = objType;
             m_ConnPt = null;

            KmpsAppl.EventList.Add(this);
        }


        ~BaseEvent()
        {
            KmpsAppl.EventList.Remove(this);

            Unadvise();

            m_Container = null;
            m_Doc = null;
        }


        // Получить имя документа
        protected string GetDocName()
        {
            string result = string.Empty;
            if (m_Doc != null)
            {
                if(m_Doc is ksDocument2D doc2d)
                {
                    ksDocumentParam docPar = KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam) as ksDocumentParam;
                    doc2d.ksGetObjParam(doc2d.reference, docPar, ldefin2d.ALLPARAM);
                    result = docPar.fileName;
                }
                else
                {
                    if(m_Doc is ksDocument3D doc3d)
                    {
                        result = doc3d.fileName;
                    }
                    else
                    {
                        ksSpcDocument spcDoc = m_Doc as ksSpcDocument;
                        if (spcDoc != null)
                        {
                            ksDocumentParam docPar = KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam) as ksDocumentParam;
                            spcDoc.ksGetObjParam(spcDoc.reference, docPar, ldefin2d.ALLPARAM);
                            result = docPar.fileName;
                        }
                        else
                        {
                            if (m_Doc is ksDocumentTxt txtDoc)
                            {
                                ksTextDocumentParam docPar = KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_TextDocumentParam) as ksTextDocumentParam;
                                txtDoc.ksGetObjParam(spcDoc.reference, txtDoc, ldefin2d.ALLPARAM);
                                result = docPar.fileName;
                            }
                        }
                    }
                }
            }
            return result;
        }


        // Подписаться на получение событий
        public int Advise()
        {
            // Подписаться на получение событий
            if (m_Container != null)
            {
                if(m_Container is IConnectionPointContainer cpContainer)
                {
                    cpContainer.FindConnectionPoint(ref m_Events, out m_ConnPt);
                    if (m_ConnPt != null)
                        m_ConnPt.Advise(this, out m_Cookie);
                }
            }

            if (m_Cookie == 0)
                return 0;


             return m_Cookie;
        }


        // Отписаться от получения событий
        void Unadvise()
        {
            if (m_Cookie == 0)
                return;

            if (m_ConnPt != null)               // Подписка была
            {
                m_ConnPt.Unadvise(m_Cookie);    // Отписаться от получения событий
                m_ConnPt = null;
            }
            m_Cookie = 0;
        }


        // Отписать все события
        public static void TerminateEvents()
        {
            int count = KmpsAppl.EventList.Count;
            for (int i = 0; i < count; i++)
            {
                BaseEvent headEvent = (BaseEvent)KmpsAppl.EventList[0];
                headEvent.Disconnect();
                KmpsAppl.EventList.Remove(headEvent);
            }
        }


        // Отписать все события по GUID и документу
        public static void TerminateEvents(Type type, object doc, int objType, ksFeature obj3D)
        {
            int count = KmpsAppl.EventList.Count;
            for (int i = count - 1; i >= 0; i--)
            {
                object obj = KmpsAppl.EventList[i];
                BaseEvent evt = (BaseEvent)obj;

                if (evt != null &&
                    (evt.GetType() == type || type == null) &&
                    (doc == null || evt.m_Doc == doc) &&
                    (objType == -1 || evt.m_ObjType == objType))
                {
                    evt.Disconnect();   // В деструкторе будет удален из списка RemoveAt(pos)
                    KmpsAppl.EventList.Remove(evt);
                }
            }
        }


        // Освободить ссылки
        void Clear()
        {
            if (m_Container != null)
            {
                m_Container = null;
            }

            if (m_Doc != null)
            {
                m_Doc = null;
            }

            m_Events = Guid.Empty;
        }


        // Отсоединиться
        void Disconnect()
        {
            Unadvise();
            Clear();
        }


        public static bool FindEvent(Type type, object doc, int objType)
        {
            int count = KmpsAppl.EventList.Count;
            for (int i = count - 1; i >= 0; i--)
            {
                object obj = KmpsAppl.EventList[i];
                BaseEvent evt = (BaseEvent)obj;
                if (evt != null &&
                    evt.GetType() == type &&
                    (doc == null || evt.m_Doc == doc) &&
                    ((objType == -1 && evt.m_ObjType == 0) || (evt.m_ObjType == objType)))
                    return true;
            }
            return false;
        }

        public static void ListEvents()
        {
            string str = string.Format("Подписанные события:");

            int count = KmpsAppl.EventList.Count;
            for (int i = count - 1; i >= 0; i--)
            {
                object obj = KmpsAppl.EventList[i];
            }
            KmpsAppl.KompasAPI.ksMessage(str);
        }
    }
}
