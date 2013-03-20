using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp;

namespace Westmark
{
    public class ITextPDF : Document
    {
        ~ITextPDF()
        {
            try
            {
                doc.Close();
                
            }
            catch {}
            doc = null;
        }
        string PathToPDF = string.Empty;
        PDFPageEvent pdfEvents;
        public ITextPDF(string FullPathToPDF, ePageSizes PageSize)
        {
            doc = new Document();
            PathToPDF = FullPathToPDF;
            writer = iTextSharp.text.pdf.PdfWriter.GetInstance(
                doc, new System.IO.FileStream(FullPathToPDF, System.IO.FileMode.Create));
            pdfEvents = new PDFPageEvent();
            writer.PageEvent = pdfEvents;

            doc.SetMargins(30, 30, 30, 30);
            switch (PageSize)
            {
                case ePageSizes.Letter:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.LETTER.Width,
                        iTextSharp.text.PageSize.LETTER.Height));
                    break;
                case ePageSizes.Letter_Landscape:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.LETTER.Height,
                        iTextSharp.text.PageSize.LETTER.Width));
                    break;
                case ePageSizes.Legal:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.LEGAL.Width,
                        iTextSharp.text.PageSize.LEGAL.Height));
                    break;
                case ePageSizes.Tabloid:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.TABLOID.Width,
                        iTextSharp.text.PageSize.TABLOID.Height));
                    break;
                case ePageSizes.Tabloid_Landscape:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.TABLOID.Height,
                        iTextSharp.text.PageSize.TABLOID.Width));
                    break;
                default:
                    doc.SetPageSize(new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.LETTER.Width,
                        iTextSharp.text.PageSize.LETTER.Height));
                    break;
            }
        }

        // **NOTE** For version >= 5.1.0.0 of iTextSharp
        Document doc = null;
        iTextSharp.text.pdf.PdfWriter writer = null;

        public enum ePageSizes
        {
            Letter,
            Letter_Landscape,
            Legal,
            Tabloid,
            Tabloid_Landscape
        }

        public enum eAlignHorizontal
        {
            Left,
            Center,
            Right
        }
        public enum eAlignVertical
        {
            Top,
            Center,
            Bottom
        }

        #region Header

        public void AddHeader(string HeaderText)
        {
            Font fnt = GetFont(eFontName.Helvitica, 10, false, false, false);
            AddHeader(HeaderText, fnt, eAlignHorizontal.Center, false, false);
        }
        public void AddHeader(string HeaderText, Font myfont)
        {
            AddHeader(HeaderText, myfont, eAlignHorizontal.Center, false, false);
        }
        public void AddHeader(string HeaderText, Font myfont, eAlignHorizontal alignment)
        {
            AddHeader(HeaderText, myfont, alignment, false, false);
        }
        public void AddHeader(string HeaderText, Font myfont, eAlignHorizontal alignment, bool IncludePageNumber)
        {
            AddHeader(HeaderText, myfont, alignment, IncludePageNumber, false);
        }
        public void AddHeader(string HeaderText, Font myfont, eAlignHorizontal alignment, bool IncludePageNumber, bool PageNumberBeforeText)
        {
            int align = 0;
            switch (alignment)
            {
                case eAlignHorizontal.Left:
                    align = Element.ALIGN_LEFT;
                    break;
                case eAlignHorizontal.Center:
                    align = Element.ALIGN_CENTER;
                    break;
                case eAlignHorizontal.Right:
                    align = Element.ALIGN_RIGHT;
                    break;
            }
            PDFPageEvent.ePageing pg = PDFPageEvent.ePageing.None;
            if (IncludePageNumber)
            {
                if (PageNumberBeforeText)
                {
                    pg = PDFPageEvent.ePageing.PageX_Before;
                }
                else
                {
                    pg = PDFPageEvent.ePageing.PageX;
                }
            }
            pdfEvents.SetHeader(HeaderText, myfont, align, pg);
        }

        #endregion

        #region Footer

        public void AddFooter(string FooterText)
        {
            Font fnt = FontFactory.GetFont(FontFactory.HELVETICA);
            AddFooter(FooterText, fnt, eAlignHorizontal.Center, false, false);
        }
        public void AddFooter(string FooterText, Font myfont)
        {
            AddFooter(FooterText, myfont, eAlignHorizontal.Center, false, false);
        }
        public void AddFooter(string FooterText, Font myfont, eAlignHorizontal alignment)
        {
            AddFooter(FooterText, myfont, alignment, false, false);
        }
        public void AddFooter(string FooterText, Font myfont, eAlignHorizontal alignment, bool IncludePageNumber)
        {
            AddFooter(FooterText, myfont, alignment, IncludePageNumber, false);
        }
        public void AddFooter(string FooterText, Font myfont, eAlignHorizontal alignment, bool IncludePageNumber, bool PageNumberBeforeText)
        {
            int align = 0;
            switch (alignment)
            {
                case eAlignHorizontal.Left:
                    align = Element.ALIGN_LEFT;
                    break;
                case eAlignHorizontal.Center:
                    align = Element.ALIGN_CENTER;
                    break;
                case eAlignHorizontal.Right:
                    align = Element.ALIGN_RIGHT;
                    break;
            }
            PDFPageEvent.ePageing pg = PDFPageEvent.ePageing.None;
            if (IncludePageNumber)
            {
                if (PageNumberBeforeText)
                {
                    pg = PDFPageEvent.ePageing.PageX_Before;
                }
                else
                {
                    pg = PDFPageEvent.ePageing.PageX;
                }
            }
            pdfEvents.SetFooter(FooterText, myfont, align, pg);
        }

        #endregion

        public void AddText(string Text, Font font, eAlignHorizontal align)
        {
            if (!doc.IsOpen()) doc.Open();
            Paragraph par = new Paragraph(Text, font);
            switch (align)
            {
                case eAlignHorizontal.Left:
                    par.Alignment = Element.ALIGN_LEFT;
                    break;
                case eAlignHorizontal.Center:
                    par.Alignment = Element.ALIGN_CENTER;
                    break;
                case eAlignHorizontal.Right:
                    par.Alignment = Element.ALIGN_RIGHT;
                    break;
            }
            doc.Add(par);
        }
        public void AddList(System.Collections.Specialized.StringCollection StringList, Font font, eAlignHorizontal align, bool WithBullet)
        {
            if (StringList.Count > 0)
            {
                iTextSharp.text.List list = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
                list.SetListSymbol("\u2022");
                list.IndentationLeft = 20f;
                for (int i = 0; i < StringList.Count; i++)
                {
                    list.Add(new iTextSharp.text.ListItem(12f, StringList[i], font));
                }

                doc.Add(list);
            }
        }
        public void AddSeparator()
        {
            iTextSharp.text.pdf.draw.LineSeparator ls = new iTextSharp.text.pdf.draw.LineSeparator();
            doc.Add(new Chunk(ls));
        }

        internal void Save(bool Close, bool OpenForView)
        {
            writer.CloseStream = true;
            // writer.Dispose();
            doc.Close();
            // doc.Dispose();
            System.Diagnostics.Process.Start(PathToPDF);
        }

        internal void AddBlank(int NumberOfBlankLines)
        {
            for (int i = 0; i < NumberOfBlankLines; i++)
            {
                doc.Add(new Paragraph(" "));
            }
        }
        public enum eFontName
        {
            Helvitica,
            Courier
        }
        public Font GetFont(eFontName FontName, float size, bool IsBold, bool IsItalic, bool IsUnderline)
        {
            Font fnt = null;
            if (IsBold)
            {
                if (IsItalic)
                {
                    switch (FontName)
                    {
                        case eFontName.Courier:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_BOLDOBLIQUE, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_BOLDOBLIQUE, size);
                            }
                            break;
                        default:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_BOLDOBLIQUE, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_BOLDOBLIQUE, size);
                            }
                            break;
                    }
                    
                }
                else
                {
                    switch (FontName)
                    {
                        case eFontName.Courier:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_BOLD, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_BOLD, size);
                            }
                            break;
                        default:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, size);
                            }
                            break;
                    }
                }
            }
            else
            {
                if (IsItalic)
                {
                    switch (FontName)
                    {
                        case eFontName.Courier:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_OBLIQUE, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER_OBLIQUE, size);
                            }
                            break;
                        default:

                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, size);
                            }
                            break;
                    }
                }
                else
                {
                    switch (FontName)
                    {
                        case eFontName.Courier:
                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.COURIER, size);
                            }
                            break;
                        default:

                            if (IsUnderline)
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA, size, Font.UNDERLINE);
                            }
                            else
                            {
                                fnt = FontFactory.GetFont(FontFactory.HELVETICA, size);
                            }
                            break;
                    }
                }
            }
            return fnt;
        }
    }

    public class PDFPageEvent : iTextSharp.text.pdf.PdfPageEventHelper
    {
        // This is the contentbyte object of the writer
        iTextSharp.text.pdf.PdfContentByte cb;

        // we will put the final number of pages in a template
        iTextSharp.text.pdf.PdfTemplate template;

        // this is the BaseFont we are going to use for the header / footer
        iTextSharp.text.pdf.BaseFont bf = null;

        // This keeps track of the creation time
        DateTime PrintTime = DateTime.Now;

        public override void OnOpenDocument(iTextSharp.text.pdf.PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);
            try
            {
                PrintTime = DateTime.Now;
                bf = FootFont.BaseFont;
                cb = writer.DirectContent;
                template = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException ex)
            {
            }
        }
        public override void OnStartPage(iTextSharp.text.pdf.PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
 
            switch (HeaderPage)
            {
                case ePageing.PageX:
                    Header = new Paragraph(HeaderText + " Page " + writer.CurrentPageNumber.ToString(), HeadFont);
                    Header.Alignment = HeaderAlign;
                    break;
                case ePageing.PageX_Before:
                    Header = new Paragraph("Page " + writer.CurrentPageNumber.ToString() +
                        " " + HeaderText, HeadFont);
                    Header.Alignment = HeaderAlign;
                    break;
                default:
                    Header = new Paragraph(HeaderText, HeadFont);
                    Header.Alignment = HeaderAlign;
                    break;
            }
 
            document.Add(Header);
    }
        
        public override void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            int pageN = writer.PageNumber;
            string text = "Page " + pageN.ToString() + " of ";
            float len = FootFont.BaseFont.GetWidthPoint(text, FootFont.Size);
            if (FooterPage != ePageing.None)
            {
                switch (FooterPage)
                {
                    case ePageing.PageX:
                        text = FooterText + "Page " + pageN.ToString() + " of ";
                        Footer.Alignment = FooterAlign;
                        break;
                    case ePageing.PageX_Before:
                        text = text = FooterText + "Page " + pageN.ToString() + " of ";
                        Footer.Alignment = FooterAlign;
                        break;
                }
                // Infinite loop - adding this to the page triggers an OnEndPage
                // document.Add(Footer);

            }

            Rectangle pageSize = document.PageSize;
            if (cb == null) cb = new iTextSharp.text.pdf.PdfContentByte(writer);
            cb.BeginText();
            cb.SetFontAndSize(FootFont.BaseFont, FootFont.Size);
            cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetBottom(30));
            cb.ShowText(text);
            cb.EndText();

            cb.AddTemplate(template, pageSize.GetLeft(40) + len, pageSize.GetBottom(30));

            cb.BeginText();
            cb.SetFontAndSize(FootFont.BaseFont, FootFont.Size);
            cb.ShowTextAligned(iTextSharp.text.pdf.PdfContentByte.ALIGN_RIGHT,
                "Printed on " + PrintTime.ToString("M/d/yy"),
                pageSize.GetRight(40),
                pageSize.GetBottom(30),
                0);
            cb.EndText();


        }

        public Paragraph Header = new Paragraph();
        public Paragraph Footer = new Paragraph();
        private Font HeadFont = FontFactory.GetFont(FontFactory.HELVETICA); // new myFont(myFont.eFontName.Helvitica);
        private Font FootFont = FontFactory.GetFont(FontFactory.HELVETICA); // new myFont(myFont.eFontName.Helvitica);
        private int HeaderAlign = Element.ALIGN_CENTER;
        private int FooterAlign = Element.ALIGN_CENTER;
        private ePageing HeaderPage = ePageing.None;
        private ePageing FooterPage = ePageing.None;
        private string HeaderText = string.Empty;
        private string FooterText = string.Empty;

        public enum ePageing
        {
            None,
            PageX,
            PageX_Before
        }
        public void SetHeader(string Content, Font HeaderFont, int Alignment, ePageing Paging)
        {
            HeaderText = Content;
            HeadFont = HeaderFont;
            HeaderAlign = Alignment;
            HeaderPage = Paging;
        }
        public void SetFooter(string Content, Font FooterFont, int Alignment, ePageing Paging)
        {
            FooterText = Content;
            FootFont = FooterFont;
            FooterAlign = Alignment;
            FooterPage = Paging;
        }

        public override void OnCloseDocument(iTextSharp.text.pdf.PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
            template.BeginText();
            template.SetFontAndSize(FootFont.BaseFont, FootFont.Size);
            template.SetTextMatrix(0, 0);
            template.ShowText("" + (writer.PageNumber - 1).ToString());
            template.EndText();
        }
    }

    

}
