using Kompas6API7;
using KompasLib.Tools;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KompasLib.KompasTool
{
    public class NotifyVariable : INotifyPropertyChanged
    {
        private IVariable7 _variable7;

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public string Name
        {
            get => this._variable7.Name;
        }

        public double Value
        {
            get => this._variable7.Value;
            set
            {
                this._variable7.Value = value;
                OnPropertyChanged("Value");
            }
        }

        public string Note
        {
            get => this._variable7.Note;
            set
            {
                this._variable7.Note = value;
                OnPropertyChanged("Note");
            }
        }

        public NotifyVariable(KmpsDoc Doc, string Name, string Index, double DefValue, string DefNote, bool Exteral = false)
        {
            if (Doc.D71.IsVariableNameValid(Name + Index))
            {
                this._variable7 = Doc.D71.AddVariable(Name + Index, DefValue, DefNote);
            }
            else
            {
                this._variable7 = Doc.D71.Variable[Exteral, Name + Index];
            }
        }

        public void Delete()
        {
            this._variable7.Delete();
        }
    }
}
