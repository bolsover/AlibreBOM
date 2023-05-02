using System.Drawing;
using Microsoft.Win32;

namespace AlibreBOM
{
    public class Globals
    {
        private static readonly string InstallPath = (string) Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Alibre Design Add-Ons\",
            "{66ae992c-e834-11ed-a05b-0242ac120003}", null);

        public static Icon Icon = new Icon(InstallPath + "\\3DPrint.ico");
        public static readonly string AppName = "Export Open Add-On ";

    }
}