using System;
using System.Globalization;
using System.Windows.Data;

namespace easyWusaWSUS.GUI
{
    /// <summary>
    /// Konverter, der prüft, ob zwei Werte gleich sind und ein Boolean zurückgibt
    /// Wird für die bedingte Aktivierung/Deaktivierung von UI-Elementen verwendet
    /// </summary>
    public class EqualityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Prüft, ob der Wert und der Parameter gleich sind
            try
            {
                // Für den Fall, dass der Parameter ein String mit einer Zahl ist
                // und der Wert ein Integer, versuchen wir eine Konvertierung
                if (parameter is string paramString && int.TryParse(paramString, out int paramInt))
                {
                    if (value is int intValue)
                    {
                        return intValue == paramInt;
                    }
                }

                return value != null && value.Equals(parameter);
            }
            catch
            {
                return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Für diesen Anwendungsfall ist keine Rückkonvertierung erforderlich
            return Binding.DoNothing;
        }
    }
}