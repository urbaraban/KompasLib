////////////////////////////////////////////////////////////////////////////////
//
// Object2DEvent - обработчик событий объектов 2D документа
//
////////////////////////////////////////////////////////////////////////////////

//#define __LIGHT_VERSION__
using Kompas6API5;
using Kompas6API7;
using KompasLib.Tools;

namespace KompasLib.Event
{
    public class Object2DEvent : BaseEvent, ksObject2DNotify
    {
        private ksObject2DNotifyResult m_res;
        public Object2DEvent(object obj, object doc, int objType,
            ksObject2DNotifyResult res, bool selfAdvise)
            : base(obj, typeof(ksObject2DNotify).GUID, doc,
            objType)
        { m_res = res; }


        // kdChangeActive - Переключение вида/слоя в текущий
        public bool ChangeActive(int viewout)
        {
            return true;
        }


        // koBeginDelete - Попытка удаления объекта
        public bool BeginDelete(int objout)
        {
            object obj = KmpsAppl.KompasAPI.TransferReference(objout, KmpsAppl.Doc.D5.reference);
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
            KmpsAppl.Doc.D5.ksLightObj(objout, 0);
            return true;
        }


        // koDelete - Объект удален
        public bool Delete(int objout)
        {
            return true;
        }


        // koBeginMove - Начало смещения объекта
        public bool BeginMove(int objout)
        {
            return true;
        }


        // koMove - Объект смещен
        public bool Move(int objout)
        {
            return true;
        }


        // koBeginRotate - Поворот объекта
        public bool BeginRotate(int objout)
        {
            return true;
        }


        // koRotate - Поворот объекта
        public bool Rotate(int objout)
        {
            return true;
        }


        // koBeginScale - Маштабирование объекта
        public bool BeginScale(int objout)
        {
            return true;
        }


        // koScale - Маштабирование объекта
        public bool scale(int objout)
        {
            return true;
        }


        // koBeginTransform - Трансформация объекта
        public bool BeginTransform(int objout)
        {
            return true;
        }


        // koTransform - Трансформация объекта
        public bool Transform(int objout)
        {
            return true;
        }


        // koBeginCopy - Копирование объекта
        public bool BeginCopy(int objout)
        {
            return true;
        }


        // koCopy - Копирование объекта
        public bool copy(int objout)
        {
            return true;
        }


        // koBeginSymmetry - Симметрия объекта
        public bool BeginSymmetry(int objout)
        {
            return true;
        }


        // koSymmetry - Симметрия объекта
        public bool Symmetry(int objout)
        {

            return true;
        }


        // koBeginProcess - Начало редактирования\создания объекта
        public bool BeginProcess(int pType, int objout)
        {
            return true;
        }


        // koEndProcess - Конец редактирования\создания объекта
        public bool EndProcess(int pType)
        {
            return true;
        }


        // koCreate - Создание объектов
        public bool CreateObject(int objout)
        {
            if (KmpsAppl.someFlag)
            {
                try
                {
                    IDrawingObject pDrawObj = (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(objout, KmpsAppl.Doc.D5.reference);
                    if (pDrawObj != null)
                    {
                        long type = (int)pDrawObj.DrawingObjectType;
                        switch (type)
                        {
                            // Линия
                            case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLineSeg:
                                {
                                    ILineSegment obj = (ILineSegment)pDrawObj;
                                    if ((obj.Style == 1) || (obj.Style == 7)) KmpsAppl.Doc.ST.SizeMe(objout, true);
                                    break;
                                }
                            // Квадрат
                            case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRectangle:
                                {
                                    IRectangle obj = (IRectangle)pDrawObj;
                                    if ((obj.Style == 1) || (obj.Style == 7)) KmpsAppl.Doc.ST.SizeMe(objout, true);
                                    break;
                                }
                        }
                    }
                }
                catch { }
            }
            return true;
        }


        // koUpdateObject - Редактирование объекта
        public bool UpdateObject(int objout)
        {
            return true;
        }
        // koUpdateObject - Редактирование объекта
        public bool BeginDestroyObject(int objout)
        {
            return true;
        }
        // koUpdateObject - Редактирование объекта
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

