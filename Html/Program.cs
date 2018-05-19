using System;
using System.Collections.Generic;
using System.Windows.Forms;
/* Project: HtmlDocumentFactory
 * Author: MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight)
 * License: BSD 3-Clause
 */

namespace Html
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
