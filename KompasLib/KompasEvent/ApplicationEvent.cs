////////////////////////////////////////////////////////////////////////////////
//
// ApplicationEvent  - ���������� ������� �� ����������
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


        // koApplicatinDestroy - �������� ����������
        public bool ApplicationDestroy()
        {
            // ������������
            TerminateEvents();
            DisconnectDoc?.Invoke(this, null);
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

        // koCreateDocument - �������� ������
        public bool CreateDocument(object newDoc, int docType)
        {
            DocumentCreated?.Invoke(newDoc, docType);
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
            DocumentChanged?.Invoke(newDoc, docType);
            return true;
        }


        // koOpenDocumen - �������� ������
        public bool OpenDocument(object newDoc, int docType)
        {
            DocumentOpened?.Invoke(newDoc, docType);
            return true;
        }

        // koKeyDown - ������� ����������
        public bool KeyDown(ref int key, int flags, bool system)
        {
            KeyboardDown?.Invoke(key, flags, system);
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
