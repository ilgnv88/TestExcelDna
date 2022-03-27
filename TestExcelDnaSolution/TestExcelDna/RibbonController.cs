using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
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
            MessageBox.Show("Hello from control " + control.Id);
            DataWriter.WriteData();
        }
    }

}
