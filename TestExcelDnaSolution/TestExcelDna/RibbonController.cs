using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using TestExcelDna.Properties;

namespace TestExcelDna
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        //public override string GetCustomUI(string ribbonId)
        //{
        //    return @"
        //      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
        //      <ribbon>
        //        <tabs>
        //          <tab id='tab1' label='My Tab'>
        //            <group id='group1' label='My Group'>
        //              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
        //            </group >
        //          </tab>
        //        </tabs>
        //      </ribbon>
        //    </customUI>";
        //}
       

        public Bitmap MyLoadImage(IRibbonControl control)
        {
            System.Drawing.Icon icon = Properties.Resources.check;
            switch (control.Id)
            {
                //case "RugbyImageButton": return icon.ToBitmap();
                //case "button1": return TestExcelDna.Properties.Resources.check1;
                case "button1": return icon.ToBitmap();
                default: return null;
            }
        }


        public void OnButtonPressed(IRibbonControl control)
        {
            Application xlApp = ExcelDnaUtil.Application as Application;

            if (xlApp == null) return;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            Range headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (int i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            Range dataRange = ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
            xlApp.ActiveChart.SetSourceData(Source: dataRange);
        }
    }

}
