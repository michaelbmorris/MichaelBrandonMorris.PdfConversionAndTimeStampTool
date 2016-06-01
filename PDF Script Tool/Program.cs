//-----------------------------------------------------------------------------------------------------------
// <copyright file="Script.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    internal static class Program
    {
        [System.STAThread]
        private static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application
                .SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new PdfScriptTool());
        }
    }
}