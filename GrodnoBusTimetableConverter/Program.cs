using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;

namespace GrodnoBusTimetableConverter
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файл с расписанием";
            ofd.Filter = "Расписание в формате xls|*.xls";
            ofd.InitialDirectory = @"D:\Files\Other\Projects\PublicTransport\Converters\GrodnoBusTimetable";//AppDomain.CurrentDomain.BaseDirectory;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Converter.Convert(ofd.FileName);
            }
        }
    }
    
}
