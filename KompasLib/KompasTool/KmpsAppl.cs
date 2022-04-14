using KAPITypes;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KompasLib.Event;
using Microsoft.Win32;
using System;
using System.Collections;
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

        public static IApplication Appl { get; set; }
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

        public static KompasObject KompasAPI { get; set; }
        #endregion

        public ApplicationEvent AppEvent { get; set; }


        [STAThread]
        public bool Connect() 
        {
            if (KmpsAppl.KompasAPI == null)
            {

                string progId = string.Empty;
#if __LIGHT_VERSION__
				progId = "KOMPASLT.Application.5";
#else
                progId = "KOMPAS.Application.5";
#endif
                try
                {
                    KmpsAppl.KompasAPI = (KompasObject)Marshal.GetActiveObject(progId);
                }
                catch (Exception)
                {
                    MessageBox.Show("Не найден активный объект", "Сообщение", MessageBoxButton.OK);
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
                            AppEvent = new ApplicationEvent(KmpsAppl.KompasAPI);
                            AppEvent.Advise();
                        }
                    }
                    catch (InvalidCastException e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
            }
            return KmpsAppl.KompasAPI != null;
        }


        public static void ZoomAll()
        {
            if (KmpsAppl.KompasAPI != null)
            {
                IDocumentFrame documentFrame = Appl.ActiveDocument.DocumentFrames[0];
                documentFrame.ZoomPrevNextOrAll(ZoomTypeEnum.ksZoomAll);
            }
        }


        public bool CreateDoc()
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
                        return doc2d.ksCreateDocument(docPar);
                    }
                }
            }
            return false;
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

        #region Реализаця интерфейса IDisposable
        public void Dispose()
        {
            if (KompasAPI != null)
            {
                Marshal.ReleaseComObject(KompasAPI);
                GC.SuppressFinalize(KompasAPI);
                KompasAPI = null;
            }
        }
        #endregion

        #region COM Registration
        // Эта функция выполняется при регистрации класса для COM
        // Она добавляет в ветку реестра компонента раздел Kompas_Library,
        // который сигнализирует о том, что класс является приложением Компас,
        // а также заменяет имя InprocServer32 на полное, с указанием пути.
        // Все это делается для того, чтобы иметь возможность подключить
        // библиотеку на вкладке ActiveX.
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas_Library");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }


        // Эта функция удаляет раздел Kompas_Library из реестра
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            subKey.DeleteSubKey("Kompas_Library");
            subKey.Close();
        }
        #endregion

    }
}
