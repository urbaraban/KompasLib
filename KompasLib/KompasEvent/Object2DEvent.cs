////////////////////////////////////////////////////////////////////////////////
//
// Object2DEvent - ���������� ������� �������� 2D ���������
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__
#if (__LIGHT_VERSION__)
	using Kompas6LTAPI5;
#else
	using Kompas6API5;
#endif
using Kompas6API7;
using KompasLib.Tools;

namespace KompasLib.Event
{
    public class Object2DEvent : BaseEvent, ksObject2DNotify
    {
        private ksObject2DNotifyResult m_res;
        private KmpsAppl api;
        public Object2DEvent(object obj, object doc, int objType,
            ksObject2DNotifyResult res, bool selfAdvise, KmpsAppl API)
            : base(obj, typeof(ksObject2DNotify).GUID, doc,
            objType)
        { m_res = res; api = API; }


        // kdChangeActive - ������������ ����/���� � �������
        public bool ChangeActive(int viewout)
        {
            return true;
        }


        // koBeginDelete - ������� �������� �������
        public bool BeginDelete(int objout)
        {
            object obj = KmpsAppl.KompasAPI.TransferReference(objout, api.Doc.D5.reference);
            try
            {
                api.Doc.D5.ksLightObj(objout, 1);
                IDrawingGroup pDrawObj = (IDrawingGroup)obj;
                if (pDrawObj != null)
                {
                    api.Doc.VisibleLayer(88, true);
                }
            }
            catch { }
            api.Doc.D5.ksLightObj(objout, 0);
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
            if (api.someFlag)
            {
                IDrawingObject pDrawObj = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(objout, api.Doc.D5.reference);
                if (pDrawObj != null)
                {
                    long type = (int)pDrawObj.DrawingObjectType;
                    switch (type)
                    {
                        // �����
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLineSeg:
                            {
                                ILineSegment obj = (ILineSegment)pDrawObj;
                                if ((obj.Style == 1) || (obj.Style == 7)) api.Doc.ST.SizeMe(objout, true);
                                break;
                            }
                        // �������
                        case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRectangle:
                            {
                                IRectangle obj = (IRectangle)pDrawObj;
                                if ((obj.Style == 1) || (obj.Style == 7)) api.Doc.ST.SizeMe(objout, true);
                                break;
                            }
                    }
                }
            }
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

