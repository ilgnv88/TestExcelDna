using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;

namespace TestExcelDna
{
    public class RibbonController : ExcelRibbon
    {
        [ComVisible(true)]
        public override string GetCustomUI(string ribbonId)
        {
            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon>
                <tabs>
                  <tab id='tab1' label='My Tab'>
                    <group id='group1' label='My Group'>
                      <button id='button1' label='My Button' onAction='OnButtonPressed'/>
                    </group >
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
            DataWriter.WriteData();
        }
    }

}
