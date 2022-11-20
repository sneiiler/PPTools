using Microsoft.Office.Tools.Ribbon;
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

        // 无用的参考代码
        //System.Windows.Forms.MessageBox.Show("ppSelectionText");
        //sel.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 10;
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

        private void line_spacing_12_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.TextRange.Count > 0)
                {
                    sel.TextRange.ParagraphFormat.SpaceWithin = 1.2f;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)//调整选中文字的行间距
                {
                    sel.TextRange.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;
                    sel.TextRange.ParagraphFormat.SpaceWithin = 1.2f;

                }
            }
            catch (Exception)
            {

                //throw;
            }
            
        }

        private void line_spacing_specific_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                float linespacing_num = float.Parse(editBox1.Text);
                if (sel.TextRange.Count > 0)
                {
                    sel.TextRange.ParagraphFormat.SpaceWithin = linespacing_num;
                }
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)//调整选中文字的行间距
                {
                    sel.TextRange.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;//True 设置行倍数;False 设置磅数
                    sel.TextRange.ParagraphFormat.SpaceWithin = linespacing_num;
                }
            }
            catch (Exception)
            {

                //throw;
            }
        }

        private void delete_current_page_animation_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Slide current_slide = app.ActiveWindow.View.Slide;

                int effect_num = int.Parse(current_slide.TimeLine.MainSequence.Count.ToString());


                for (int i = current_slide.TimeLine.MainSequence.Count; i >= 1; i--)
                {
                    Effect effect = current_slide.TimeLine.MainSequence[i];
                    effect.Delete();
                }
                MessageBox.Show("已删除" + effect_num + "个动画效果!", "动画删除结果");
            }
            catch (Exception)
            {

                //throw;
            }
        }

        // 删除选中页面所有动画
        private void delete_selected_page_animation_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selected_slides = app.ActiveWindow.Selection.SlideRange;

                foreach (PowerPoint.Slide current_slides in selected_slides)
                {
                    for (int i = current_slides.TimeLine.MainSequence.Count; i >= 1; i--)
                    {
                        Effect effect = current_slides.TimeLine.MainSequence[i];
                        effect.Delete();
                    }
                }
            }
            catch (Exception)
            {

                //throw;
            }
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
             
        }

        private void font_weiruanyahei_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                sel.TextRange.Font.NameFarEast = "微软雅黑";
            }
            catch (Exception)
            {

                //throw;
            }

        }

        private void font_timesNR_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                sel.TextRange.Font.NameOther = "Times New Roman";
                sel.TextRange.Font.Name = "Times New Roman";
            }
            catch (Exception)
            {

                //throw;
            }

        }

        private void resize_shape_fullbackground_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var slide_height = app.ActivePresentation.PageSetup.SlideHeight;// 以磅为单位返回当前幻灯片尺寸，一磅约为0.3527778毫米
                var slide_width = app.ActivePresentation.PageSetup.SlideWidth;// 以磅为单位返回当前幻灯片尺寸，一磅约为0.3527778毫米
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes) 
                { 
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    foreach (PowerPoint.Shape shape in range) 
                    {
                        if (isLockAspectRatio.Checked == false)
                        {
                            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                        }
                        else
                        {
                            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                        }
                        shape.Width = slide_width; 
                        shape.Height = slide_height;
                        shape.Left = 0;
                        shape.Top = 0;
                    
                    }
                }
            }
            catch (Exception)
            {

                //throw;
            }
            
        }
    }
}
