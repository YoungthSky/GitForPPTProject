using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PPTGenerator
{
    public class PPTOperater
    {
        #region ppt操作对象信息
        POWERPOINT.Application m_PptApp;//PPT的操作程序
        POWERPOINT.Presentation m_PptPresSet;//当前打开的PPT
        //POWERPOINT.SlideShowSettings m_PptSSS;
        //bool bAssistantOn;

        POWERPOINT.Slide m_CurSlide = null;        //PPT中的幻灯片

        private string m_CurTargetFilePath = null;//当前打开PPT的路径
        private int m_SlideCount = 1;
        #endregion

        #region 文件打开保存关闭操作
        public void TTPOpen(string filePath)
        {
            //防止连续打开多个PPT程序.
            if (this.m_PptApp != null) { return; }
            try
            {
                m_PptApp = new POWERPOINT.Application();
                //以非只读方式打开,方便操作结束后保存.
                m_PptPresSet = m_PptApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoCTrue);
                m_CurTargetFilePath = filePath;
            }
            catch (Exception ex)
            {
                this.m_PptApp.Quit();
                m_PptApp = null;
            }
        }
        public void PPTCreate(string filePath)
        {
            string folderPath = filePath.Substring(0, filePath.LastIndexOf('/'));
            if(!Directory.Exists(folderPath))
            {
                return;
            }
            m_PptApp = new POWERPOINT.Application();
            m_PptPresSet = m_PptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
            m_CurTargetFilePath = filePath;
            //bAssistantOn = m_PptApp.Assistant.On;
            //m_PptApp.Assistant.On = false;
            if(File.Exists(m_CurTargetFilePath))
            {
                File.Delete(m_CurTargetFilePath);
            }
        }

        public void PPTSave(string filePath)
        {
            try
            {
                filePath = filePath.Replace('/', '\\');
                if (filePath.Equals(m_PptPresSet.FullName))
                {
                    m_PptPresSet.Save();
                }
                else
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    POWERPOINT.PpSaveAsFileType format = POWERPOINT.PpSaveAsFileType.ppSaveAsDefault;
                    m_PptPresSet.SaveAs(filePath, format, Microsoft.Office.Core.MsoTriState.msoFalse);
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        public void PPTClose(bool isSave)
        {
            if(m_PptPresSet != null)
            {
                if(isSave && m_CurTargetFilePath != null)
                    PPTSave(m_CurTargetFilePath);
                m_PptPresSet.Close();
            }
            if (m_PptApp != null)
                m_PptApp.Quit();
            GC.Collect();
        }
        #endregion

        public void AddSlide(POWERPOINT.PpSlideLayout layout)
        {
            AddSlide(m_SlideCount++, layout);
        }
        public void AddSlide(int index, POWERPOINT.PpSlideLayout layout)
        {
            if(m_PptPresSet!=null)
                m_CurSlide = m_PptPresSet.Slides.Add(index, layout);
        }

        public void AddWords(string tittle, string[] words)
        {
            POWERPOINT.TextRange myTextRng = null;

            if(m_CurSlide == null)
            {
                AddSlide(POWERPOINT.PpSlideLayout.ppLayoutTitleOnly);
            }
            POWERPOINT.Shape ts = m_CurSlide.Shapes.Title;
            if (ts != null && ts.Type == MsoShapeType.msoTextBox)
            {
                myTextRng = ts.TextFrame.TextRange;
                myTextRng.Font.NameFarEast = "微软雅黑";
                myTextRng.Font.NameAscii = "Calibri";
                myTextRng.Text = tittle;
                myTextRng.Font.Bold = MsoTriState.msoCTrue;
                myTextRng.Font.Color.RGB = 0 + 0 * 256 + 0 * 256 * 256;
                myTextRng.Characters(1, 30).Font.Size = 48;
                myTextRng.ParagraphFormat.Alignment = POWERPOINT.PpParagraphAlignment.ppAlignCenter;
                ts.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            }
            for (int i = 0; i < words.Length; ++i)
            {
                POWERPOINT.Shape s = m_CurSlide.Shapes.AddTextbox(OFFICECORE.MsoTextOrientation.msoTextOrientationHorizontal, 20f, 120f + i * 50f, 500f, 50f);
                if (s!= null && s.Type == MsoShapeType.msoTextBox)
                {
                    myTextRng = s.TextFrame.TextRange;
                    myTextRng.Font.NameFarEast = "微软雅黑";
                    myTextRng.Font.NameAscii = "Calibri";
                    myTextRng.Text = words[i];
                    myTextRng.Font.Bold = MsoTriState.msoCTrue;
                    myTextRng.Font.Color.RGB = 100 + 0 * 256 + 0 * 256 * 256;
                    myTextRng.Characters(1, 30).Font.Size = 24;
                    myTextRng.ParagraphFormat.Alignment = POWERPOINT.PpParagraphAlignment.ppAlignLeft;
                    s.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                }
            }
        }
        /// <summary>
        /// 添加一页ppt并且设置页面标题和内容记录
        /// </summary>
        /// <param name="tittle">标题</param>
        /// <param name="words">内容记录</param>
        public void AddSlide(string tittle, string[] words)
        {
            AddSlide(POWERPOINT.PpSlideLayout.ppLayoutTitleOnly);
            AddWords(tittle, words);
        }
        private bool GoToSlide(int index)
        {
            if (m_PptPresSet != null && index <= m_PptPresSet.Slides.Count)
            {
                m_CurSlide = m_PptPresSet.Slides[index];
                return true;
            }
            return false;
        }

        /// <summary>
        /// 修改对应页面标题
        /// </summary>
        /// <param name="sindex">页面索引</param>
        /// <param name="word">要设置的标题内容</param>
        /// <returns>设置成功与否</returns>
        public bool ChangeTittle(int sindex, string word)
        {
            if (sindex <= 0)
            {
                return false;
            }
            if (GoToSlide(sindex))
            {
                try
                {
                    if (m_CurSlide.Shapes.Title.Type == MsoShapeType.msoTextBox)
                    {
                        POWERPOINT.TextRange myTextRng = m_CurSlide.Shapes.Title.TextFrame.TextRange;
                        myTextRng.Text = word;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }
        /// <summary>
        /// 修改对应页面内容
        /// </summary>
        /// <param name="sindex">页面索引</param>
        /// <param name="windex">记录值索引</param>
        /// <param name="word">修改的记录内容</param>
        /// <returns>设置成功与否</returns>
        public bool ChangeContent(int sindex, int windex, string word)
        {
            if(sindex <= 0 || windex <= 0)
            {
                return false;
            }
            if(GoToSlide(sindex))
            {
                try
                {
                    if (windex < m_CurSlide.Shapes.Count && m_CurSlide.Shapes[windex + 1].Type == MsoShapeType.msoTextBox)
                    {
                        POWERPOINT.TextRange myTextRng = m_CurSlide.Shapes[windex + 1].TextFrame.TextRange;
                        myTextRng.Text = word;
                        return true;
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }
        /// <summary>
        /// 修改对应页面第windex个文本框记录内容，包括了页面标题在windex索引内
        /// </summary>
        /// <param name="sindex">页面索引</param>
        /// <param name="windex">记录值索引</param>
        /// <param name="word">修改的记录内容</param>
        /// <returns>设置成功与否</returns>
        public bool ChangeWordRecord(int sindex, int windex, string word)
        {
            if (sindex <= 0 || windex <= 0)
            {
                return false;
            }
            if (GoToSlide(sindex))
            {
                try
                {
                    if (windex <= m_CurSlide.Shapes.Count && m_CurSlide.Shapes[windex].Type == MsoShapeType.msoTextBox)
                    {
                        POWERPOINT.TextRange myTextRng = m_CurSlide.Shapes[windex].TextFrame.TextRange;
                        myTextRng.Text = word;
                        return true;
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }

        public void AddImg(string imgpath, float x, float y, float width, float heigh)
        {
            AddSlide(POWERPOINT.PpSlideLayout.ppLayoutVerticalTitleAndText);
            m_CurSlide.Shapes.AddPicture(imgpath, MsoTriState.msoFalse, MsoTriState.msoCTrue, x, y, width, heigh);
        }
        public bool InsertImg(int sindex, string imgpath,float x,float y,float width, float heigh)
        {
            try
            {
                if (GoToSlide(sindex))
                {
                    m_CurSlide.Shapes.AddPicture(imgpath, MsoTriState.msoFalse, MsoTriState.msoCTrue, x, y, width, heigh);
                    return true;
                }
            }
            catch
            {
                return false;
            }
            return false;
        }

        /// <summary>
        /// 替换幻灯片的首张图片
        /// </summary>
        /// <param name="sindex">幻灯片索引</param>
        /// <param name="imgpath">新的图片路径</param>
        /// <returns>替换是否成功返回</returns>
        public bool ExchangeImg(int sindex, string imgpath, MsoZOrderCmd layoutType = MsoZOrderCmd.msoSendToBack)
        {
            try
            {
                if (GoToSlide(sindex))
                {
                    for (int i = 1; i <= m_CurSlide.Shapes.Count; ++i)
                    {
                        POWERPOINT.Shape s = m_CurSlide.Shapes[i];
                        if (s != null && s.Type == MsoShapeType.msoPicture)
                        {
                            //POWERPOINT.TextFrame pic = s.Width;
                            float width = s.Width;
                            float left = s.Left;
                            float top = s.Top;
                            float height = s.Height;
                            s.Delete();
                            s = m_CurSlide.Shapes.AddPicture(imgpath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, width, height);
                            s.ZOrder(layoutType);
                        }
                    }
                    return true;
                }
            }
            catch
            {
                return false;
            }
            return false;
        }

        /// <summary>
        /// 更改幻灯片的第imgdndex张图片
        /// </summary>
        /// <param name="sindex">幻灯片索引</param>
        /// <param name="imgIndex">图片索引</param>
        /// <param name="imgpath">新的图片路径</param>
        /// <returns>替换是否成功返回</returns>
        public bool ExchangeImg(int sindex, int imgIndex, string imgpath, MsoZOrderCmd layoutType = MsoZOrderCmd.msoSendToBack)
        {
            try
            {
                if (GoToSlide(sindex))
                {
                    List<POWERPOINT.Shape> _shapList = new List<POWERPOINT.Shape>();
                    for(int i= 1; i <= m_CurSlide.Shapes.Count; ++i)
                    {
                        POWERPOINT.Shape s = m_CurSlide.Shapes[i];
                        if (s != null && s.Type == MsoShapeType.msoPicture)
                        {
                            _shapList.Add(s);
                        }
                    }
                    if (_shapList.Count >= imgIndex && imgIndex > 0)
                    {
                        POWERPOINT.Shape s = _shapList[imgIndex-1];
                        if (s != null && s.Type == MsoShapeType.msoPicture)
                        {
                            //POWERPOINT.TextFrame pic = s.Width;
                            float width = s.Width;
                            float left = s.Left;
                            float top = s.Top;
                            float height = s.Height;
                            s.Delete();
                            s = m_CurSlide.Shapes.AddPicture(imgpath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, width, height);
                            s.ZOrder(layoutType);
                        }
                    }
                    return true;
                }
            }
            catch
            {
                return false;
            }
            return false;
        }

    }
}
