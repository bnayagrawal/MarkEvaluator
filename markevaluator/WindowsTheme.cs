using System;
using Microsoft.Win32;
using System.Drawing;

namespace markevaluator
{
    class WindowsTheme
    {
        /// <summary>
        /// Extracts windows accent color value from windows registry
        /// Tested on windows 10, may not work in other versions.
        /// </summary>
        /// <returns>hex string value</returns>
        public static Color getColorFromRegistry()
        {
            Color c;
            try
            {
                //Extract Windows color theme from registry
                Object val = Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\DWM", "ColorizationColor", Color.FromArgb(46, 170, 197, 1));
                string hexValue = "#" + ((Int32)val).ToString("X");
                c = (Color)(new ColorConverter()).ConvertFromString(hexValue);
            }
            catch(Exception ex)
            {
                c = Color.FromArgb(46,170,197,1);
                LogWriter.WriteError("While fetching accent color from registry", ex.Message);
            }
            return c;
        }

        /// <summary>
        /// Returns Hex color value from registry
        /// </summary>
        /// <returns></returns>
        public static String getHexColorFromRegistry()
        {
            Color c = getColorFromRegistry();
            return ("#FF" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2"));
        }

        /// <summary>
        /// Returns a darker shade Hex color value from registry
        /// </summary>
        /// <returns></returns>
        public static String getHexDarkColorFromRegistry()
        {
            Color c = getColorFromRegistry();
            c = Color.FromArgb(c.A, (int)(c.R * 0.8), (int)(c.G * 0.8), (int)(c.B * 0.8));
            return ("#FF" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2"));
        }
    }
}
