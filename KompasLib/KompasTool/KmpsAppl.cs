using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Event;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;

namespace KompasLib.Tools
{
    public class KmpsAppl : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        #region Variable
        //
        //private
        //

        private static  ksMathematic2D mat;
        private  IPropertyManager propMng;
        private static IProgressBarIndicator progressBar;
        private KmpsDoc doc;

        //
        //public
        //

        //Изменили документ
        public event EventHandler CreatedDoc;
        public event EventHandler<object> CreatedObject;
        public event EventHandler<object> SelectedObject;
        public event EventHandler<Tuple<int, bool>> EventKeyUp;

        //Отключились
        public  event EventHandler<bool> ConnectBoolEvent;

        public KmpsDoc Doc
        {
            get => this.doc;
            set
            {
                this.doc = value;
                OnPropertyChanged("Doc");
            }
        }

        public static  bool someFlag = true;

        public static  IApplication Appl { get; set; }
        public static ksMathematic2D Mat
        {
            get => mat;
        }

         public static  IProgressBarIndicator ProgressBar
        {
            get => Appl.ProgressBarIndicator;
            set => progressBar = value;
        }

        public  IPropertyManager PropMng
        {
            get => propMng;
            set => propMng = value;
        }

        public static ArrayList EventList { get; set; } = new ArrayList();

        public static  KompasObject KompasAPI { get; set; }
        #endregion

        public KmpsAppl()
        {

        }

        public bool Connect() 
        {
            if (KmpsAppl.KompasAPI == null)
            {
                string progId = "KOMPAS.Application.5";
                try
                {
                    KmpsAppl.KompasAPI = (KompasObject)Marshal.GetActiveObject(progId);
                }
                catch (InvalidCastException e)
                {
                    MessageBox.Show(e.Message);
                    return false;
                }

                if (KmpsAppl.KompasAPI != null)
                {
                    try
                    {
                        KmpsAppl.KompasAPI.Visible = true;
                        KmpsAppl.KompasAPI.ActivateControllerAPI();

                        Appl = (IApplication)KmpsAppl.KompasAPI.ksGetApplication7();
                        if (Appl == null)
                            return false;

                        mat = (ksMathematic2D)KmpsAppl.KompasAPI.GetMathematic2D();

                        ConnectBoolEvent?.Invoke(null, true);

                        if (!BaseEvent.FindEvent(typeof(ApplicationEvent), null, -1))
                        {
                            ApplicationEvent aplEvent = new ApplicationEvent(KmpsAppl.KompasAPI);
                            aplEvent.Advise();
                            aplEvent.OpenedDoc += AplEvent_OpenedDoc;
                            aplEvent.CreatedDoc += AplEvent_CreatedDoc;
                            aplEvent.ChangeDoc += AplEvent_ChangeDoc;
                            aplEvent.EvKeyUp += AplEvent_EvKeyUp;
                        }

                        SelectDoc();
                    }
                    catch (InvalidCastException e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
            }
            return KmpsAppl.KompasAPI != null;
        }

        private  void AplEvent_EvKeyUp(object sender, Tuple<int, bool> e)
        {
            EventKeyUp?.Invoke(null, e);
        }

        private  void AplEvent_ChangeDoc(object sender, object e)
        {
            SelectDoc();
        }

        private  void AplEvent_CreatedDoc(object sender, object e)
        {
            SelectDoc();
        }

        private  void AplEvent_OpenedDoc(object sender, object e)
        {
            SelectDoc();
        }

        public  bool OpenFile(string filepath)
        {
            int type = KmpsAppl.KompasAPI.ksGetDocumentTypeByName(filepath);
            switch (type)
            {

                case (int)DocType.lt_DocSheetStandart:  //2d документы
                case (int)DocType.lt_DocFragment:
                    this.Doc.D5 = (ksDocument2D)KmpsAppl.KompasAPI.Document2D();
                    if (this.Doc.D5 != null)
                        Process.Start(filepath);
                    return true;

            }

            int err = KmpsAppl.KompasAPI.ksReturnResult();
            if (err != 0)
                KmpsAppl.KompasAPI.ksResultNULL();
            return false;
        }
        public static void ZoomAll()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IDocumentFrame documentFrame = Appl.ActiveDocument.DocumentFrames[0];
                documentFrame.ZoomPrevNextOrAll(ZoomTypeEnum.ksZoomAll);
            }
        }

        public  void SelectDoc()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                KmpsDoc tempDoc = new KmpsDoc(this);

                if (tempDoc.D5 != null)
                {
                    int[] YesType = { 1, 4, 9, 35, 27 };

                    for (int i = 0; i < YesType.Length; i++)
                    {
                        if (!BaseEvent.FindEvent(typeof(DocumentEvent), tempDoc, YesType[i]))
                            AdviseDoc((Kompas6API5.ksDocumentFileNotify_Event)KmpsAppl.KompasAPI.ksGetDocumentByReference(tempDoc.D5.reference), KmpsAppl.KompasAPI.ksGetDocumentType(tempDoc.D5.reference), YesType[i]);
                    }

                    this.Doc = tempDoc;
                }             
            }
        }

        public  bool CreateDoc()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                ksDocument2D doc2d = (ksDocument2D)KmpsAppl.KompasAPI.Document2D();
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
                    ksDocumentParam docPar = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

                    if (docPar != null)
                    {
                        docPar.Init();
                        docPar.regime = 0;
                        docPar.type = (int)DocType.lt_DocFragment;
                        docPar.author = Environment.UserName;
                        docPar.comment = "KPCm";
                        if (doc2d.ksCreateDocument(docPar))
                        {
                            return true;
                        }
                        else return false;
                    }
                }
            }
            return false;
        }


        private  void AdviseDoc(object doc, int docType, int objType/*-1*/)
        {
            if (doc == null)
                return;

            // События документа, необходимы для своевременной отписки
            if (!BaseEvent.FindEvent(typeof(DocumentEvent), doc, objType))
            {
                // Обработчик событий от документа
                DocumentEvent docEvent = new DocumentEvent((Kompas6API5.ksDocumentFileNotify_Event)doc);
                // Подписка на события документа
                int advise = docEvent.Advise();

                // Неудачная подписка на события документа
                if (advise == 0)
                    return;
            }
            else
                KmpsAppl.KompasAPI.ksError("На события документа уже подписались");

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
                                Document2DEvent document2DEvent = new Document2DEvent(doc2DNotify, doc2D);
                                document2DEvent.Advise();
                            }
                        }


                    // Селектирование

                    if (!BaseEvent.FindEvent(typeof(SelectMngEvent), doc2D, objType))
                    {
                        object selMsg = doc2D?.GetSelectionMngNotify();
                        if (selMsg != null)
                        {
                            SelectMngEvent selEvent = new SelectMngEvent(selMsg, doc2D);
                            selEvent.Advise();
                            selEvent.SelectedObject += SelEvent_SelectedObject;
                        }
                    }


                    // Объект 2D документа
                    if (objType >= 0) // Тип приходит всегда
                    {
                        if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc2D, objType))
                        {
                            object objNotify = doc2D?.GetObject2DNotify(objType);
                            if (objNotify != null)
                            {
                                Object2DEvent objEvent = new Object2DEvent(objNotify, doc2D, objType, doc2D.GetObject2DNotifyResult());
                                objEvent.Advise();
                                objEvent.OnCreatedObjectRef += ObjEvent_OnCreatedObjectRef;
                            }
                        }
                    }
                    break;
            }
        }

        private  void SelEvent_SelectedObject(object sender, int e)
        {
            SelectedObject?.Invoke(null, (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(e, this.Doc.D5.reference));
        }

        private  void ObjEvent_OnCreatedObjectRef(object sender, int e)
        {
            if (KmpsAppl.someFlag)
            {
                try
                {
                    CreatedObject?.Invoke(null, (IDrawingObject)KmpsAppl.KompasAPI.TransferReference(e, this.Doc.D5.reference));
                }
                catch { }
            }
        }

        public  void DisconnectKmps()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                // принудительно зпрервать св¤зь с  омпас
                Marshal.ReleaseComObject(KmpsAppl.KompasAPI);
                ConnectBoolEvent?.Invoke(null, false);
            }
        }

    }
}
