////////////////////////////////////////////////////////////////////////////////
//
// Object2DEvent - ���������� ������� �������� 2D ���������
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__
using Kompas6API5;
using System;

namespace KompasLib.Event
{
    public class Object2DEvent : BaseEvent, ksObject2DNotify
    {
        public event EventHandler<int> OnCreatedObjectRef;

        private ksObject2DNotifyResult m_res;
        public Object2DEvent(object obj, object doc, int objType,
            ksObject2DNotifyResult res)
            : base(obj, typeof(ksObject2DNotify).GUID, doc,
            objType)
        { m_res = res; }


        // kdChangeActive - ������������ ����/���� � �������
        public bool ChangeActive(int viewout)
        {
            return true;
        }


        // koBeginDelete - ������� �������� �������
        public bool BeginDelete(int objout)
        {
           /* object obj = KmpsAppl.KompasAPI.TransferReference(objout, KmpsAppl.Doc.D5.reference);
            try
            {
                KmpsAppl.Doc.D5.ksLightObj(objout, 1);
                IDrawingGroup pDrawObj = (IDrawingGroup)obj;
                if (pDrawObj != null)
                {
                    KmpsAppl.Doc.VisibleLayer(88, true);
                }
            }
            catch { }
            KmpsAppl.Doc.D5.ksLightObj(objout, 0);*/
            return true;
        }


        // koDelete - ������ ������
        public bool Delete(int objout)
        {
            return true;
        }


        // koBeginMove - ������ �������� �������
        public bool BeginMove(int objout)
        {
            return true;
        }


        // koMove - ������ ������
        public bool Move(int objout)
        {
            return true;
        }


        // koBeginRotate - ������� �������
        public bool BeginRotate(int objout)
        {
            return true;
        }


        // koRotate - ������� �������
        public bool Rotate(int objout)
        {
            return true;
        }


        // koBeginScale - �������������� �������
        public bool BeginScale(int objout)
        {
            return true;
        }


        // koScale - �������������� �������
        public bool scale(int objout)
        {
            return true;
        }


        // koBeginTransform - ������������� �������
        public bool BeginTransform(int objout)
        {
            return true;
        }


        // koTransform - ������������� �������
        public bool Transform(int objout)
        {
            return true;
        }


        // koBeginCopy - ����������� �������
        public bool BeginCopy(int objout)
        {
            return true;
        }


        // koCopy - ����������� �������
        public bool copy(int objout)
        {
            return true;
        }


        // koBeginSymmetry - ��������� �������
        public bool BeginSymmetry(int objout)
        {
            return true;
        }


        // koSymmetry - ��������� �������
        public bool Symmetry(int objout)
        {
            return true;
        }


        // koBeginProcess - ������ ��������������\�������� �������
        public bool BeginProcess(int pType, int objout)
        {
            return true;
        }


        // koEndProcess - ����� ��������������\�������� �������
        public bool EndProcess(int pType)
        {
            return true;
        }


        // koCreate - �������� ��������
        public bool CreateObject(int objout)
        {
            OnCreatedObjectRef(this, objout);
            return true;
        }


        // koUpdateObject - �������������� �������
        public bool UpdateObject(int objout)
        {
            return true;
        }
        // koUpdateObject - �������������� �������
        public bool BeginDestroyObject(int objout)
        {
            return true;
        }
        // koUpdateObject - �������������� �������
        public bool DestroyObject(int objout)
        {
            return true;
        }
        public bool BeginPropertyChanged(int objout)
        {
            return true;
        }

        public bool PropertyChanged(int objout)
        {
            return true;
        }
    }
}

