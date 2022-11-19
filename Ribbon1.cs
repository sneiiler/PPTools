using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace PPTools
{
    public partial class Ribbon1
    {
        PowerPoint.Application app;// 实例化当前PPT
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)  //加载事件
        {
            app= Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            sel.TextRange.ParagraphFormat.SpaceWithin = 1.2f;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText) {//调整选中文字的行间距
                //System.Windows.Forms.MessageBox.Show("ppSelectionText");
                //sel.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 10;
                sel.TextRange.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;
                sel.TextRange.ParagraphFormat.SpaceWithin = 1.2f;
                //PowerPoint.TextRange range= sel.TextRange;
                //foreach (PowerPoint.TextRange shape in range)
                //{
                //    //shape.ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStyleMixed;
                //    //shape.Width= shape.Width*2;
                //    //shape.Height= shape.Height*2;
                //    //if (shape.Type == Office.MsoShapeType.msoTextBox)
                //    //{
                //    //    shape.TextEffect.FontSize = shape.TextEffect.FontSize + 1; 
                //    //}

                //    shape.Text = "fdfd0";



                //}
            }
        }
    }
}
