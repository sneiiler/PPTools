﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

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

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            float linespacing_num = float.Parse(editBox1.Text);
            sel.TextRange.ParagraphFormat.SpaceWithin = linespacing_num;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)//调整选中文字的行间距
            {
                sel.TextRange.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;//True 设置行倍数;False 设置磅数
                sel.TextRange.ParagraphFormat.SpaceWithin = linespacing_num;
            }


           
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Slide current_slide = app.ActiveWindow.View.Slide;

            int effect_num = int.Parse(current_slide.TimeLine.MainSequence.Count.ToString());


            for (int i = current_slide.TimeLine.MainSequence.Count; i >= 1; i--)
            {
                Effect effect = current_slide.TimeLine.MainSequence[i];
                effect.Delete();
                //for (int x = sequence.Count; x >= 1; x--)
                //{
                //    Effect effect = sequence[x];
                    
                //}
            }
            MessageBox.Show("已删除" + effect_num + "个动画效果!", "动画删除结果");
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //PowerPoint.SlideRange selected_slides = app.ActiveWindow.View.s;


            //foreach (Slide current_slide in selected_slides)
            //{
            //    for (int i = current_slide.TimeLine.MainSequence.Count; i >= 1; i--)
            //    {
            //        Effect effect = current_slide.TimeLine.MainSequence[i];
            //        effect.Delete();
            //        //for (int x = sequence.Count; x >= 1; x--)
            //        //{
            //        //    Effect effect = sequence[x];

            //        //}
            //    }
            //}

        }
    }
}
