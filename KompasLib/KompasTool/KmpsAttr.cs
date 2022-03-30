using KAPITypes;
using Kompas6API5;
using Kompas6Constants;
using reference = System.Int32;

namespace KompasLib.Tools
{
    public class KmpsAttr
    {
        private KmpsDoc kmpsDoc;
        private ksAttributeObject attributeObject;

        public KmpsAttr(KmpsDoc kmps)
        {
            this.kmpsDoc = kmps;
            attributeObject = (ksAttributeObject)KmpsAppl.KompasAPI.GetAttributeObject();
        }

        public ksAttributeObject AO => attributeObject;

        public reference NewAttr(reference pObj)
        {
            ksAttributeParam attrPar = (ksAttributeParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_Attribute);
            ksUserParam usPar = (ksUserParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_UserParam);
            ksDynamicArray fVisibl = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(23);
            ksDynamicArray colKeys = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(23);
            if (attrPar != null && usPar != null && fVisibl != null && colKeys != null)
            {
                attrPar.Init();
                usPar.Init();
                attrPar.SetValues(usPar);
                attrPar.SetColumnKeys(colKeys);
                attrPar.SetFlagVisible(fVisibl);
                attrPar.key1 = 1;
                attrPar.key2 = 10;
                attrPar.key3 = 100;
                attrPar.password = string.Empty;

                ksLtVariant item = (ksLtVariant)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                ksDynamicArray arr = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(23);
                if (item != null && arr != null)
                {
                    usPar.SetUserArray(arr);
                    item.Init();
                    item.floatVal = 1;
                    arr.ksAddArrayItem(-1, item);
                    item.Init();
                    item.floatVal = 2;
                    arr.ksAddArrayItem(-1, item);
                    item.Init();
                    item.floatVal = 3;
                    arr.ksAddArrayItem(-1, item);
                    item.Init();
                    item.floatVal = 4;
                    arr.ksAddArrayItem(-1, item);
                    item.Init();
                    item.floatVal = 5;
                    arr.ksAddArrayItem(-1, item);
                    item.Init();
                    item.floatVal = 6;
                    arr.ksAddArrayItem(-1, item);
                }

                if (this.kmpsDoc.D5.ksExistObj(pObj) != 0)
                {
                    double numb = GiveIDNameTypeAttr("ForMacroParam");
                    return attributeObject.ksCreateAttr(pObj, attrPar, numb, null);
                }
            }
            return 0;
        }

        public void FuncAttrType()
        {
            ksAttributeTypeParam type = (ksAttributeTypeParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
            ksColumnInfoParam col = (ksColumnInfoParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);
            if (type != null && col != null)
            {
                type.Init();
                col.Init();
                type.header = "ForMacroParam";    // заголовoк-комментарий типа
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
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // координата Y снизу
                    col.header = "mathBop.Y";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // Координата Х сверху
                    col.header = "mathTop.X";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // координата Y сверху
                    col.header = "mathTop.Y";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // размер по Х
                    col.header = "sizeX";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
                    col.key = 0;                            // дополнительный признак, который позволит отличить две переменные с одинаковым типом
                    col.def = "0";                          // значение по умолчанию
                    col.flagEnum = false;                   // флаг включающий режим, когда значение поля атрибута
                    arr.ksAddArrayItem(-1, col);

                    // размер Y
                    col.header = "sizeY";                  // заголовoк-комментарий столбца
                    col.type = ldefin2d.FLOAT_ATTR_TYPE;   // тип данных в столбце - см.ниже
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

        public reference GiveObjAttr(reference pObj)
        {

            //создадим итератор для хождения по атрибутам объекта
            ksIterator iter = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
            if (iter != null && iter.ksCreateAttrIterator(pObj, 0, 0, 0, 0, 0))
            {
                //встали на первый атрибут
                reference pAttr = iter.ksMoveAttrIterator("F", ref pObj);
                if (pAttr != 0)
                    return pAttr;
                else
                    KmpsAppl.KompasAPI.ksMessage("атрибут не найден");
            }
            return 0;
        }
    }
}
