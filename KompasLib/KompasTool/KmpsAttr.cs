using KAPITypes;
using Kompas6API5;
using Kompas6Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KompasLib.Tools
{
    public class KmpsAttr
    {
        private KmpsAppl api;
        private ksAttributeObject attributeObject;

        public KmpsAttr(KmpsAppl API)
        {
            api = API;
            attributeObject = (ksAttributeObject)KmpsAppl.KompasAPI.GetAttributeObject();
        }

        public ksAttributeObject AO
        {
            get => attributeObject;
        }

        public void FuncAttrType()
        {

            ksAttributeTypeParam type = (ksAttributeTypeParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
            ksColumnInfoParam col = (ksColumnInfoParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
            if (type != null && col != null)
            {
                type.Init();
                col.Init();
                type.header = "ForCoordinaMacro";    // заголовoк-комментарий типа
                type.rowsCount = 1;                 // кол-во строк в таблице
                type.flagVisible = true;            // видимый, невидимый   в таблице
                type.password = string.Empty;       // пароль, если не пустая строка  - защищает от несанкционированного изменения типа
                type.key1 = 10;
                type.key2 = 20;
                type.key3 = 30;
                type.key4 = 0;
                ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
                if (arr != null)
                {
                    /*attribute.AddRow(string.Empty, 0);
                    attribute.SetValue(string.Empty, 0, 0, );
                    attribute.SetValue(string.Empty, 0, 1, mathBop.y);
                    attribute.SetValue(string.Empty, 0, 2, mathTop.x);
                    attribute.SetValue(string.Empty, 0, 3, mathTop.y);
                    attribute.SetValue(string.Empty, 0, 4, sizeX);
                    attribute.SetValue(string.Empty, 0, 5, sizeY);*/

                    // Координата Х снизу
                    col.header = "mathBop.X";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // координата Y снизу
                    col.header = "mathBop.Y";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // Координата Х сверху
                    col.header = "mathTop.X";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // координата Y сверху
                    col.header = "mathTop.Y";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // размер по Х
                    col.header = "sizeX";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // размер Y
                    col.header = "sizeY";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.DOUBLE_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);
                }
                string nameFile = string.Empty;
                //запросить имя библиотеки

                //создать тип атрибута
                double numbType = 0;
                if (GiveIDNameTypeAttr(type.header) == 0)
                    numbType = attributeObject.ksCreateAttrType(type,   // информация о типе атрибута
                    null);                                  // имя библиотеки типов атрибутов
                //удалим  массив колонок
                arr.ksDeleteArray();
            }
        }


        private double GiveIDNameTypeAttr(string name)
        {
            //Получаем типы
            ksDynamicArray DynamicArray = (ksDynamicArray)attributeObject.ksGetLibraryAttrTypesArray(null);
            int count = DynamicArray.ksGetArrayCount();
            //перебираем
            for (int i = 0; i < count; i++)
            {
                ksLibraryAttrTypeParam typeParam = (ksLibraryAttrTypeParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_LibraryAttrTypeParam);
                DynamicArray.ksGetArrayItem(i, typeParam);
                if (typeParam.name == name) return typeParam.typeId; //Если такое имя есть, то возвращаем айди
            }
            return 0;
        }
    }
}
