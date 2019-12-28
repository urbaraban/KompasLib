using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Event;
using System;
using System.Collections;
using System.Runtime.InteropServices;


namespace KompasLib.Tools
{
    public class KmpsAppl
    {
        #region Variable
        //
        //private
        //

        private static KompasObject kompasAPI;
        private static IApplication appl;
        private ksMathematic2D mat;
        private IPropertyManager propMng;
        private IProgressBarIndicator progressBar;
        private KmpsDoc doc;

        private bool closeDocChek;
        private int closeDocValue;
        


        //
        //public
        //


        //Изменили документ
        public event EventHandler<KmpsAppl> ChangeDoc;
        //Отключились
        public event EventHandler<bool> ConnectBoolEvent;

        public KmpsDoc Doc
        {
            get => doc;
            set => doc = value;
        }



        public bool someFlag = true;




        public static IApplication Appl
        {
            get => appl;
            set => appl = value;
        }

        public ksMathematic2D Mat
        {
            get => this.mat = (ksMathematic2D)kompasAPI.GetMathematic2D();
        }

         public IProgressBarIndicator ProgressBar
        {
            get => appl.ProgressBarIndicator;
            set => progressBar = value;
        }

        public IPropertyManager PropMng
        {
            get => propMng;
            set => propMng = value;
        }

        private static ArrayList eventList = new ArrayList();
        public static ArrayList EventList
        {
            get => eventList;
            set => eventList = value;
        }

        public static KompasObject KompasAPI
        {
            get => kompasAPI;
            set => kompasAPI = value;
        }
        #endregion

        public KmpsAppl(int CloseDocValue, bool CloseDocChek)
        {
            if (kompasAPI == null)
            {
                string progId = "KOMPAS.Application.5";
                try
                {
                    kompasAPI = (KompasObject)Marshal.GetActiveObject(progId);
                }
                catch { }

                if (kompasAPI != null)
                {
                    kompasAPI.Visible = true;
                    kompasAPI.ActivateControllerAPI();
                    
                    appl = (IApplication)kompasAPI.ksGetApplication7();
                    if (appl == null)
                        return;

                    if (ConnectBoolEvent != null)
                    ConnectBoolEvent(this, true);

                    if (!BaseEvent.FindEvent(typeof(ApplicationEvent), null, -1))
                    {
                        ApplicationEvent aplEvent = new ApplicationEvent(kompasAPI, this);
                        aplEvent.Advise();
                    }
                }
            }
        }


        public void SelectDoc()
        {
            if (kompasAPI != null)
            {
                this.doc = new KmpsDoc(this);
                if (doc.D5 != null)
                {
                    if (ChangeDoc != null)
                        ChangeDoc(this, this);

                    int[] YesType = { 1, 4, 9, 35, 27 };

                    for (int i = 0; i < YesType.Length; i++)
                    {
                        if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc, YesType[i]))
                            AdviseDoc((Kompas6API5.ksDocumentFileNotify_Event)kompasAPI.ksGetDocumentByReference(doc.D5.reference), kompasAPI.ksGetDocumentType(doc.D5.reference), YesType[i]);
                    }
                }
            }
        }

        public bool CreateDoc()
        {
            ksDocument2D doc2d = (ksDocument2D)kompasAPI.Document2D();
            // создать новый документ
            // первый параметр - тип открываемого файла
            //  0 - лист чертежа
            //  1 - фрагмент
            //  2 - текстовый документ
            //  3 - спецификация
            //  4 - 3D-модель
            // второй параметр указывает на необходимость выдачи запроса "Файл изменен. Сохранять?" при закрытии файла
            // третий параметр - указатель на IDispatch, по которому Графие вызывает уведомления об изенении своего состояния
            // ф-ия возвращает HANDLE открытого документа
            if (doc2d != null)
            {
                ksDocumentParam docPar = (ksDocumentParam)kompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

                if (docPar != null)
                {
                    docPar.Init();
                    docPar.regime = 0;
                    docPar.type = (int)DocType.lt_DocFragment;
                    docPar.author = Environment.UserName;
                    docPar.comment = "KPCm";
                    if (doc2d.ksCreateDocument(docPar))
                        return true;
                    else return false;
                }
            }
            return false;
        }


        private void AdviseDoc(object doc, int docType, int objType/*-1*/)
        {
            if (doc == null)
                return;

            // События документа, необходимы для своевременной отписки
            if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc, objType))
            {
                // Обработчик событий от документа
                DocumentEvent docEvent = new DocumentEvent((Kompas6API5.ksDocumentFileNotify_Event)doc, this);
                // Подписка на события документа
                int advise = docEvent.Advise();

                // Неудачная подписка на события документа
                if (advise == 0)
                    return;
            }
            else
                kompasAPI.ksError("На события документа уже подписались");

            switch (docType)
            {
                case (int)DocType.lt_DocSheetStandart:      // 1 - чертеж стандартный
                case (int)DocType.lt_DocSheetUser:          // 2 - чертеж нестандартный
                case (int)DocType.lt_DocFragment:           // 3 - фрагмент
                    ksDocument2D doc2D = (ksDocument2D)doc; // Интерфейс документа

                    // Документ 2D
                        if (!BaseEvent.FindEvent(typeof(Document2DEvent), doc2D, objType))
                        {
                            object doc2DNotify = doc2D.GetDocument2DNotify();
                            if (doc2DNotify != null)
                            {
                                Document2DEvent document2DEvent = new Document2DEvent(doc2DNotify, doc2D, true);
                                document2DEvent.Advise();
                            }
                        }


                    // Селектирование

                        if (!BaseEvent.FindEvent(typeof(SelectMngEvent), doc2D, objType))
                        {
                            object selMsg = doc2D != null ? doc2D.GetSelectionMngNotify() : null;
                            if (selMsg != null)
                            {
                                SelectMngEvent selEvent = new SelectMngEvent(selMsg, doc2D, true);
                                selEvent.Advise();
                            }
                        }


                    // Объект 2D документа
                    if (objType >= 0) // Тип приходит всегда
                    {
                        if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc2D, objType))
                        {
                            object objNotify = doc2D != null ? doc2D.GetObject2DNotify(objType) : null;
                            if (objNotify != null)
                            {
                                Object2DEvent objEvent = new Object2DEvent(objNotify, doc2D, objType, doc2D.GetObject2DNotifyResult(), true, this);
                                objEvent.Advise();
                            }
                        }

                    }
                    break;

            }
        }


        public void DisconnectKmps()
        {
            if (kompasAPI != null)
            {
                // принудительно зпрервать св¤зь с  омпас
                Marshal.ReleaseComObject(kompasAPI);
                if (ConnectBoolEvent != null)
                    ConnectBoolEvent(this, false);
            }
        }

    }
}
