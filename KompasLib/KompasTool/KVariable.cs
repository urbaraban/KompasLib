﻿using Kompas6API7;
using KompasLib.KompasTool;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace KompasLib.Tools
{
    public class KVariable
    {
        private KmpsDoc _doc;
        public KVariable(KmpsDoc Doc)
        {
            this._doc = Doc;
        }

        //удаляем переменные
        public async Task RemoveAsync(string name, string index)
        {
            await Task.Factory.StartNew(() =>
             {
                 if (!this._doc.D71.IsVariableNameValid(name + index))
                     this._doc.D71.Variable[false, name + index].Delete();
             });

        }

        public IVariable7 Variable(string Name, string Index, bool External = false)
        {
            return this._doc.D71.Variable[External, Name + Index];
        }

        public bool IsVariableNameValid(string Name)
        {
            return this._doc.D71.IsVariableNameValid(Name);
        }

            
        //Добавляем в переменную

        public static void SetVarToUIElement(NotifyVariable variable, Control UIE, DependencyProperty dependencyProperty, IValueConverter converter, string param)
        {
            Binding binding = new Binding { Source = variable, Path = new PropertyPath(param), Mode = BindingMode.TwoWay };
            if (converter != null) binding.Converter = converter;

            UIE.SetBinding(dependencyProperty, binding);
        }
    }

    public struct VariableStruct
    {
        public string Name;
        public double Value;
        public string Note;
        public bool Base;
        public bool Info;
    }

}
