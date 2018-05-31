using Microsoft.Office.Core;
using System;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PPTGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            PPTOperater popt = new PPTOperater();
            popt.TTPOpen("D:/lyf/world.ppt");
            popt.AddImg("D:\\lyf\\mount.jpg", 0, 0, 250, 200);
            popt.AddWords("hello world",new string[] { "江南水乡美如画", "借问酒家何处有", "牧童遥指杏花村" });
            popt.AddSlide("hello china", new string[] { "塞外风光", "大漠孤烟直", "长河落日圆" });
            //popt.ChangeTittle(1, "厉害了我的锅");
            //popt.ChangeContent(2, 3, "美好世界！！");
            
            //popt.InsertImg(1,"D:\\lyf\\mount.jpg", 0, 200, 250, 200);
            popt.ExchangeImg(1,"D:\\lyf\\sea.jpg");
            popt.PPTSave("D:/lyf/world0.ppt");
            popt.PPTClose(false);
        }
            //static void Main(string[] args)
            //{
            //    string filePath = "D:/lyf/world.ppt";
            //    string folderPath = filePath.Substring(0, filePath.LastIndexOf('/'));
            //    if (!Directory.Exists(folderPath))
            //    {
            //        return;
            //    }

            //    POWERPOINT.Application PPT = new POWERPOINT.Application();//创建PPT应用
            //    POWERPOINT.Presentation MyPres = null;  //PPT应用实例
            //    POWERPOINT.Slide MySlide = null;        //PPT中的幻灯片


            //    MyPres = PPT.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);//.Open("D://lyf/hello.ppt", OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoTrue);

            //    //var x = MyPres.Slides.Application.slides;
            //    MySlide = MyPres.Slides.Add(1, POWERPOINT.PpSlideLayout.ppLayoutTitleOnly);

            //    POWERPOINT.TextRange myTextRng = null;
            //    MySlide.Shapes.AddTextbox(OFFICECORE.MsoTextOrientation.msoTextOrientationHorizontal, 21.5f, 36f, 670f, 50f);
            //    myTextRng = MySlide.Shapes[1].TextFrame.TextRange;
            //    myTextRng.Font.NameFarEast = "微软雅黑";
            //    myTextRng.Font.NameAscii = "Calibri";
            //    myTextRng.Text = "C#生成PPT00000000";
            //    myTextRng.Font.Bold = MsoTriState.msoCTrue;
            //    myTextRng.Font.Color.RGB = 100 + 0 * 256 + 150 * 256 * 256;
            //    myTextRng.Characters(1, 10).Font.Size = 24;
            //    myTextRng.ParagraphFormat.Alignment = POWERPOINT.PpParagraphAlignment.ppAlignLeft;
            //    MySlide.Shapes[1].TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

            //    MySlide.Shapes.AddPicture("http://e.hiphotos.baidu.com/image/pic/item/d6ca7bcb0a46f21fca6fafecfa246b600c33ae32.jpg", MsoTriState.msoFalse, MsoTriState.msoCTrue, 21.5F, 86F, 300F, 150F);


            //    Microsoft.Office.Interop.PowerPoint.Table MyTable = null;

            //    MyTable = MySlide.Shapes.AddTable(4, 5, 40F, 100F, 500F, 400F).Table;//创建时规定的宽和高，不是表格最终的大小。

            //    for (int k = 1; k <= MyTable.Rows.Count; ++k)
            //        for (int j = 1; j <= MyTable.Columns.Count; ++j)
            //        {
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Font.Size = 10;
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Font.Color.RGB = 250 + 250 * 256 + 100 * 256 * 256;
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Font.NameAscii = "Arial";
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Font.NameFarEast = "微软雅黑";
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            //            MyTable.Cell(k, j).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            //            MyTable.Cell(k, j).Shape.Fill.ForeColor.RGB = 0;
            //            MyTable.Cell(k, j).Shape.TextFrame.TextRange.Text = "C#生成PPTTabel";
            //        }
            //    POWERPOINT.PpSaveAsFileType format = POWERPOINT.PpSaveAsFileType.ppSaveAsDefault;
            //    MyPres.SaveAs("D://lyf/world.ppt", format, Microsoft.Office.Core.MsoTriState.msoFalse);
            //    MyPres.Close();
            //    PPT.Quit();
            //}
    }
}
