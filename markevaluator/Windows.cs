using System.Windows.Shell;
using System.Windows;

namespace markevaluator
{
    /// <summary>
    /// This class holds the refference of windows used in this program
    /// and for setting windows border thickness
    /// </summary>
    class Windows
    {
        public static admin_window adminWindow;
        public static MainWindow loginWindow;
        public static parser_window parserWindow;
        public static generator_window generatorWindow;

        /// <summary>
        /// Sets the native window border thickness to almost invisible
        /// </summary>
        /// <param name="wobj">Window object</param>
        public static void setWindowChrome(Window wobj)
        {
            WindowChrome wc = new WindowChrome();
            wc.ResizeBorderThickness = new Thickness(3);
            wc.GlassFrameThickness = new Thickness(1);
            wc.CaptionHeight = 1;
            wc.UseAeroCaptionButtons = false;
            WindowChrome.SetWindowChrome(wobj, wc);
        }
    }
}
