﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KompasLib.KompasTool
{
    public static class HlpClasses
    {
        /// <summary>
        /// Клас загрузки в ComboBox
        /// </summary>

        public class ComboBT
        {
            public string Name { get; set; }
            public string Mac { get; set; }

            public ComboBT(string _name, string _mac)
            {
                Name = _name;
                Mac = _mac;
            }
        }

        public class ComboData
        {
            public string Name { get; set; }
            public Int32 Reference { get; set; }

            public Int32 NumberPP { get; set; }

            public ComboData(string _name, Int32 _ref, Int32 _num)
            {
                Name = _name;
                Reference = _ref;
                NumberPP = _num;
            }
        }
    }
}
