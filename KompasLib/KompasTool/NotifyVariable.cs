﻿using Kompas6API7;
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
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        public string Name
        {
            get => this._variable7.Name;
        }

        public double Value
        {
            get => this._variable7 != null ? this._variable7.Value : 0;
            set
            {
                if (this._variable7 != null)
                {
                    this._variable7.Value = value;
                    OnPropertyChanged("Value");
                }
            }
        }

        public string Note
        {
            get => this._variable7 != null ? this._variable7.Note : String.Empty;
            set
            {
                if (this._variable7 != null)
                {
                    this._variable7.Note = value;
                    OnPropertyChanged("Note");
                }
            }
        }

        public NotifyVariable(KmpsDoc Doc, string Name, double DefValue, string DefNote, bool Exteral = false)
        {
            if (Doc.D71.IsVariableNameValid(Name) == true)
            {
                this._variable7 = Doc.D71.AddVariable(Name, DefValue, DefNote);
                if (this._variable7 != null)
                {
                    this._variable7.Value = DefValue;
                    this._variable7.Note = DefNote;
                }
            }
            else
            {
                this._variable7 = Doc.D71.Variable[Exteral, Name];
            }
        }

        public NotifyVariable(IVariable7 variable7)
        {
                this._variable7 = variable7;
        }

        public void Delete()
        {
            this._variable7.Delete();
        }

    }
}
