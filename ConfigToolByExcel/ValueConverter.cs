namespace ConfigToolByExcel
{
    public delegate void CustomConverter(string valueText, out Type type, out object value);

    internal class ValueConverter
    {
        private static Dictionary<string, CustomConverter> m_converters = new Dictionary<string, CustomConverter>
        {
            { "short", ShortConverter },
            { "int", IntConverter },
            { "long", LongConverter },
            { "float", FloatConverter },
            { "double", DoubleConverter },
            { "string", StringConverter },
            { "short[]", ShortArrayConverter },
            { "int[]", IntArrayConverter },
            { "long[]", LongArrayConverter },
            { "float[]", FloatArrayConverter },
            { "double[]", DoubleArrayConverter },
            { "string[]", StringArrayConverter },
        };

        private const string ArraySplitSymbol = "#";

        public static bool IsValidType(string valueTypeText)
        {
            return m_converters.ContainsKey(valueTypeText);
        }

        public static bool TryConvertValue(string valueTypeText, string valueText, out Type? type, out object? value)
        {
            if (m_converters.ContainsKey(valueTypeText))
            {
                m_converters[valueTypeText].Invoke(valueText, out type, out value);
                return true;
            }
            else
            {
                type = null;
                value = null;
                return false;
            }
        }

        private static void ShortConverter(string valueText, out Type type, out object value)
        {
            type = typeof(short);
            value = Convert.ChangeType(valueText, type);
        }

        private static void IntConverter(string valueText, out Type type, out object value)
        {
            type = typeof(int);
            value = Convert.ChangeType(valueText, type);
        }

        private static void LongConverter(string valueText, out Type type, out object value)
        {
            type = typeof(long);
            value = Convert.ChangeType(valueText, type);
        }

        private static void FloatConverter(string valueText, out Type type, out object value)
        {
            type = typeof(float);
            value = Convert.ChangeType(valueText, type);
        }

        private static void DoubleConverter(string valueText, out Type type, out object value)
        {
            type = typeof(double);
            value = Convert.ChangeType(valueText, type);
        }

        private static void StringConverter(string valueText, out Type type, out object value)
        {
            type = typeof(string);
            value = Convert.ChangeType(valueText, type);
        }

        private static void ShortArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(short[]);
            value = ParseArray(valueText, typeof(short));
        }

        private static void IntArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(int[]);
            value = ParseArray(valueText, typeof(int));
        }

        private static void LongArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(long[]);
            value = ParseArray(valueText, typeof(long));
        }

        private static void FloatArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(float[]);
            value = ParseArray(valueText, typeof(float));
        }

        private static void DoubleArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(double[]);
            value = ParseArray(valueText, typeof(double));
        }

        private static void StringArrayConverter(string valueText, out Type type, out object value)
        {
            type = typeof(string[]);
            value = ParseArray(valueText, typeof(string));
        }

        private static Array ParseArray(string arrayText, Type elementType)
        {
            var splitedElementsText = arrayText.Split(ArraySplitSymbol);
            Array array = Array.CreateInstance(elementType, splitedElementsText.Length);
            for (int i = 0; i < splitedElementsText.Length; i++)
            {
                var elementValue = Convert.ChangeType(splitedElementsText[i], elementType);
                array.SetValue(elementValue, i);
            }
            return array;
        }
    }
}
