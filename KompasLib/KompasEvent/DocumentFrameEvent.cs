using Kompas6API5;
using Kompas6API7;
using KompasLib.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KompasLib.Event
{
    public class DocumentFrameEvent : BaseEvent, ksDocumentFrameNotify
    {
        public DocumentFrameEvent(object obj, object doc, bool selfAdvise)
            : base(obj, typeof(ksDocumentFrameNotify).GUID, doc,
            -1)
        { }

        // frBeginPaint - Начало отрисовки документа
        public bool BeginPaint(IPaintObject paintObj)
        {
            return true;
        }

        // frClosePaint - Конец отрисовки документа
        public bool ClosePaint(IPaintObject paintObj)
        {
             return true;
        }

        // frMouseDown - Нажатие кнопки мыши
        public bool MouseDown(short nButton, short nShiftState, int x, int y)
        {
            return true;
        }

        // frMouseUp - Отпускание кнопки мыши
        public bool MouseUp(short nButton, short nShiftState, int x, int y)
        {
            return true;
        }

        // frMouseMove - Движение мыши
        public bool MouseMove(short nShiftState, int x, int y)
        {
            if (SizeTool.SelectPointFlag)
            {
                IDrawingObject drwObj = (IDrawingObject)SizeTool.SelectObj;
                SizeTool.PhGroup.Clear(true);
                SizeTool.PhGroup.Open();
                SizeTool.PhGroup.AddObjects((object)KmpsAppl.KompasAPI.TransferReference(KmpsAppl.Doc.D5.ksPoint(x, y, new Random(8).Next()), KmpsAppl.Doc.D5.reference));
                SizeTool.PhGroup.Close();
            }

            return true;
        }

        // frMouseDblClick - Двойной клик кнопки мыши
        public bool MouseDblClick(short nButton, short nShiftState, int x, int y)
        {
            return true;
        }

        // frBeginPaintGL - Начало отрисовки в контексте OpenGL
        public bool BeginPaintGL(ksGLObject glObj, int drawMode)
        {
            return true;
        }

        // frClosePaintGL - Окончание отрисовки в контексте OpenGL
        public bool ClosePaintGL(ksGLObject glObj, int drawMode)
        {
            return true;
        }

        // frAddGabarit - Определение габаритов документа
        public bool AddGabarit(IGabaritObject gabObj)
        {
            return true;
        }

        // frBeginCurrentProcess - Начало текущего процесса
        public bool BeginCurrentProcess(int id)
        {
            return true;
        }

        // frStopCurrentProcess - Окончание текущего процесса
        public bool StopCurrentProcess(int id)
        {
            return true;
        }

        // frActivate - Окно активизировалось
        public bool Activate()
        {

            return true;
        }

        // frDeactivate - Окно деактивизировалось
        public bool Deactivate()
        {

            return true;
        }

        // frCloseFrame - Закрытие окна
        public bool CloseFrame()
        {
            BaseEvent.TerminateEvents();
            return true;
        }

        public bool ShowOcxTree(object ocx, bool show)
        {
            return true;
        }
    }
}
