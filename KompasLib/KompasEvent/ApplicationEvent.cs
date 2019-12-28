////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - ���������� ������� �� ����������
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__

using Kompas6API5;
using Kompas6API7;
using KompasLib.Tools;
using System.Threading;

namespace KompasLib.Event
{
    public class ApplicationEvent : BaseEvent, ksKompasObjectNotify
    {
        private KmpsAppl api;
        public ApplicationEvent(object obj, KmpsAppl API)
          : base(obj, typeof(ksKompasObjectNotify).GUID,
          null, -1)
        { this.api = API; }


        // koApplicatinDestroy - �������� ����������
        public bool ApplicationDestroy()
        {
            // ������������
            TerminateEvents();
            KmpsAppl.KompasAPI = null;
            Thread thread = new Thread(api.DisconnectKmps);
            thread.Start(); 
            return true;
        }


        // koBeginCloseAllDocument - ������ �������� ���� �������� ����������
        public bool BeginCloseAllDocument()
        {
            bool res = true;
            return res;
        }


        // koBeginCreate - ������ �������� ���������(�� ������� ������ ����)
        public bool BeginCreate(int docType)
        {
            return true;
        }


        // koOpenDocumenBegin - ������ �������� ���������
        public bool BeginOpenDocument(string fileName)
        {
            return true;
        }


        // koBeginOpenFile - ������ �������� ���������(�� ������� ������ �����)
        public bool BeginOpenFile()
        {
            return true;
        }


        // koActiveDocument - ������������ �� ������ �������� ��������
        public bool ChangeActiveDocument(object newDoc, int docType)
        {
            Thread thread = new Thread(api.SelectDoc);
            thread.Start();
            return true;
        }


        // koCreateDocument - �������� ������
        public bool CreateDocument(object newDoc, int docType)
        {
            Thread thread = new Thread(api.SelectDoc);
            thread.Start();
            return true;
        }


        // koOpenDocumen - �������� ������
        public bool OpenDocument(object newDoc, int docType)
        {
            Thread thread = new Thread(api.SelectDoc);
            thread.Start();
            return true;
        }

        // koKeyDown - ������� ����������
        public bool KeyDown(ref int key, int flags, bool system)
        {
            return true;
        }

        // koKeyUp - ������� ����������
        public bool KeyUp(ref int key, int flags, bool system)
        {
            return true;
        }

        // koKeyPress - ������� ����������
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
