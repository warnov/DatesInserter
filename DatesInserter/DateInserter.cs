using System;
using Microsoft.Office.Tools.Ribbon;

namespace DatesInserter
{
    public partial class DateInserter
    {
        private static Microsoft.Office.Interop.Excel.Application 
            App => Globals.ThisAddIn.Application;
        private void btnToday_Click(object sender, RibbonControlEventArgs e)
        {
            App.ActiveCell.Value = DateTime.Today;
        }

        private void btnNextWeek_Click(object sender, RibbonControlEventArgs e)
        {
            App.ActiveCell.Value = DateTime.Today + TimeSpan.FromDays(7);
        }

        private void btnPreviousWeek_Click(object sender, RibbonControlEventArgs e)
        {
            App.ActiveCell.Value = DateTime.Today - TimeSpan.FromDays(7);
        }

        private void txtDelta_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int delta;
            if (int.TryParse(txtDelta.Text, out delta))
            {
                 App.ActiveCell.Value =  DateTime.Today + TimeSpan.FromDays(delta);
            }

        }
    }
}
