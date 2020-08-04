using System;
using System.Collections.Specialized;
using System.IO;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes;
using MigraDocCore.DocumentObjectModel.Shapes;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using PdfSharpCore.Fonts;
using PdfSharpCore.Utils;


namespace PDFGen_test01
{
    class Program
    {
        static Document document;

        static Table table;
        private static TextFrame addressFrame;
        
        static void Main(string[] args)
        {
            document = new Document();
            table = new Table();
            addressFrame = new TextFrame();
            CreateDocument();
            Console.WriteLine("Hello World!");
        }
        
        
        public static void CreateDocument()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            GlobalFontSettings.FontResolver = new FileFontResolver();
           // ImageSource.ImageSourceImpl = new ImageSharpImageSource<>();
           // ImageSource.ImageSourceImpl = new MyImageSource();
           // MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes
            ImageSource.ImageSourceImpl = new ImageSharpImageSource<SixLabors.ImageSharp.PixelFormats.Rgba32>();
            //ImageSource.ImageSourceImpl = new ImageSharpImageSource();
            
            
            // Create a new MigraDoc document
            document = new Document();
            document.Info.Title = "A sample invoice";
            document.Info.Subject = "Demonstrates how to create an invoice.";
            document.Info.Author = "Stefan Lange";
            
            
           
           // var font = new XFont(FontResolver("sans-serif", 20));
            
            Section testSection = document.AddSection();
          //   var textFrame = testSection.AddTextFrame();
          //   var paragraph = textFrame.AddParagraph();
          //  // paragraph.AddText("SOME VALUE");
          //   
          //   paragraph = testSection.AddParagraph();
          //   paragraph.Format.SpaceBefore = "8cm";
          //   paragraph.Style = "Reference";
          // //  paragraph.AddFormattedText("INVOICE", TextFormat.Bold);
          //   paragraph.AddTab();
            // paragraph.AddText("Cologne, ");
            // paragraph.AddDateField("dd.MM.yyyy");
            
          Style style = document.Styles["Normal"];
          style.Font.Name = "Verdana";
          //style.Font.Name = "Verdana";
          //style.Font.Name = "Times New Roman";
          style.Font.Size = 9;
           
        
            DefineStyles();

            CreatePage();

            string documentPath = $@"/Users/dannybolick/MainNas/PersonalRepo/GeneralTestProjectsMain/PdfSharpCoreFork/PDFGen_test01/bin/Debug/TEST.pdf";
            
         //   pdfd.Save(documentPath);
            
            FillContent();
          
           // Create a renderer for PDF that uses Unicode font encoding
           PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer();

           // Set the MigraDoc document
           pdfRenderer.Document = document;
           
         
           // Create the PDF document
           pdfRenderer.RenderDocument();
           pdfRenderer.Save(documentPath); 
        }

       static void  DefineStyles()
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";
        
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);
        
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);
        
            // Create a new style called Table based on style Normal
            style = document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Name = "Times New Roman";
            style.Font.Size = 9;
        
            // Create a new style called Reference based on style Normal
            style = document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
        }


       static void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            Section section = document.AddSection();
        
            // Put a logo in the header
            
           // section.Headers = new HeadersFooters();
         //  ImageSource source = new MyImageSource();
         string documentPath = $@"/Users/dannybolick/MainNas/PersonalRepo/GeneralTestProjectsMain/PdfSharpCoreFork/PDFGen_test01/bin/Debug/";

         var path =documentPath + "test2.png";
//         var imagepath ="/Users/dannybolick/MainNas/Personal/Projects/PersonalProjects/varible_test/PDFGen_Test/PDFGen_Test_01/bin/Debug/netcoreapp3.1/";

         MyImageSourceImp imageSourceImp = new MyImageSourceImp(path);
         
     
         document.ImagePath = documentPath;
         
         // Image img = new Image();
         //
         // section.Add(img);
         
         //var image = section.AddImage(imageSourceImp);//"/Users/dannybolick/MainNas/Personal/Projects/PersonalProjects/varible_test/PDFGen_Test/PDFGen_Test_01/bin/Debug/netcoreapp3.1/test2.png")//  
           Image image = section.Headers.Primary.AddImage(imageSourceImp);//AddImage(@"/Users/dannybolick/MainNas/Personal/Projects/PersonalProjects/varible_test/PDFGen_Test/PDFGen_Test_01/bin/Debug/netcoreapp3.1/test2.png");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Right;
            image.WrapFormat.Style = WrapStyle.Through;
        
           
           #region DATA
            // Create footer
            Paragraph paragraph = section.Footers.Primary.AddParagraph();
            paragraph.AddText("PowerBooks Inc · Sample Street 42 · 56789 Cologne · Germany");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
        
            // Create the text frame for the address
             addressFrame = section.AddTextFrame();
            addressFrame.Height = "3.0cm";
            addressFrame.Width = "7.0cm";
            addressFrame.Left = ShapePosition.Left;
            addressFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            addressFrame.Top = "5.0cm";
            addressFrame.RelativeVertical = RelativeVertical.Page;
            
            // Put sender in address frame
            paragraph = addressFrame.AddParagraph("PowerBooks Inc · Sample Street 42 · 56789 Cologne");
            paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Font.Size = 7;
            paragraph.Format.SpaceAfter = 3;
        
            // Add the print date field
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "8cm";
            paragraph.Style = "Reference";
            paragraph.AddFormattedText("INVOICE", TextFormat.Bold);
            paragraph.AddTab();
            paragraph.AddText("Cologne, ");
            paragraph.AddDateField("dd.MM.yyyy");
        
            // Create the item table
            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Colors.Red;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;
        
            // Before you can add a row, you must define the columns
            Column column = table.AddColumn("1cm");
            column.Format.Alignment = ParagraphAlignment.Center;
        
            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;
        
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;
        
            column = table.AddColumn("3.5cm");
            column.Format.Alignment = ParagraphAlignment.Right;
        
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
        
            column = table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Right;
        
            // Create the header of the table
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.Blue;
            row.Cells[0].AddParagraph("Item");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("Title and Author");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].MergeRight = 3;
            row.Cells[5].AddParagraph("Extended Price");
            row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[5].MergeDown = 1;
        
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.Blue;
            row.Cells[1].AddParagraph("Quantity");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].AddParagraph("Unit Price");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("Discount (%)");
            row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[4].AddParagraph("Taxable");
            row.Cells[4].Format.Alignment = ParagraphAlignment.Left;
        
           // table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.White);
            table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75,Colors.Aqua);
            #endregion 
        }

        static void FillContent()
        {
            // Fill address in address text frame
           // XPathNavigator item = SelectItem("/invoice/to");
            Paragraph paragraph = addressFrame.AddParagraph();
            paragraph.AddText("single_NAME");
            paragraph.AddLineBreak();
            paragraph.AddText("ADDRESS LINE !");
            paragraph.AddLineBreak();
            paragraph.AddText("CITY VALUE");
        
            // Iterate the invoice items
            double totalExtendedPrice = 0;
         //   XPathNodeIterator iter = navigator.Select("/invoice/items/*");
           // while (iter.MoveNext())
            for(int i = 0;i < 10 ; i++)
            {
               // item = iter.Current;
                double quantity = 100;//GetValueAsDouble(item, "quantity");
                double price = 2.00;//GetValueAsDouble(item, "price");
                double discount = 12;//GetValueAsDouble(item, "discount");
        
                // Each item fills two rows
                Row row1 = table.AddRow();
                Row row2 = table.AddRow();
                row1.TopPadding = 1.5;
                row1.Cells[0].Shading.Color = Colors.Gray;
                row1.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].MergeDown = 1;
                row1.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Cells[1].MergeRight = 3;
                row1.Cells[5].Shading.Color = Colors.Coral;
                row1.Cells[5].MergeDown = 1;
        
                row1.Cells[0].AddParagraph("1000");
                paragraph = row1.Cells[1].AddParagraph();
                paragraph.AddFormattedText("title_ITEM", TextFormat.Bold);
                paragraph.AddFormattedText(" by ", TextFormat.Italic);
                paragraph.AddText("AUthor_ITEM");
                row2.Cells[1].AddParagraph("100");
                row2.Cells[2].AddParagraph(price.ToString("0.00") + " €");
                row2.Cells[3].AddParagraph(discount.ToString("0.0"));
                row2.Cells[4].AddParagraph();
                row2.Cells[5].AddParagraph(price.ToString("0.00"));
                double extendedPrice = quantity * price;
                extendedPrice = extendedPrice * (100 - discount) / 100;
                row1.Cells[5].AddParagraph(extendedPrice.ToString("0.00") + " €");
                row1.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
                totalExtendedPrice += extendedPrice;
        
                table.SetEdge(0, table.Rows.Count - 2, 6, 2, Edge.Box, BorderStyle.Single, 0.75);
            }
        
            // Add an invisible row as a space line to the table
            Row row = table.AddRow();
            row.Borders.Visible = false;
        
            // Add the total price row
            row = table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("Total Price");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
        
            // Add the VAT row
            row = table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("VAT (19%)");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].AddParagraph((0.19 * totalExtendedPrice).ToString("0.00") + " €");
        
            // Add the additional fee row
            row = table.AddRow();
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].AddParagraph("Shipping and Handling");
            row.Cells[5].AddParagraph(0.ToString("0.00") + " €");
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
        
            // Add the total due row
            row = table.AddRow();
            row.Cells[0].AddParagraph("Total Due");
            row.Cells[0].Borders.Visible = false;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 4;
            totalExtendedPrice += 0.19 * totalExtendedPrice;
            row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
        
            // Set the borders of the specified cell range
            table.SetEdge(5, table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);
        
            // Add the notes paragraph
            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.SpaceBefore = "1cm";
            paragraph.Format.Borders.Width = 0.75;
            paragraph.Format.Borders.Distance = 3;
            paragraph.Format.Borders.Color = Colors.Black;
            paragraph.Format.Shading.Color = Colors.Gray;
          //  item = SelectItem("/invoice");
            paragraph.AddText("NOTES");
        }
        
    }
    
    public class FileFontResolver :  IFontResolver 
    {
        public string DefaultFontName => "Envy Code R";

        public byte[] GetFont(string faceName)
        {
            using (var ms = new MemoryStream())
            {
                using (var fs = File.Open(faceName, FileMode.Open))
                {
                    fs.CopyTo(ms);
                    ms.Position = 0;
                    return ms.ToArray();
                }
            }
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            return new FontResolverInfo("/Users/dannybolick/Library/Fonts/Envy Code R.ttf");
            if (familyName.Equals("Verdana", StringComparison.CurrentCultureIgnoreCase))
            {
                if (isBold && isItalic)
                {
                    return new FontResolverInfo("Fonts/Verdana-BoldItalic.ttf");
                }
                else if (isBold)
                {
                    return new FontResolverInfo("Fonts/Verdana-Bold.ttf");
                }
                else if (isItalic)
                {
                    return new FontResolverInfo("Fonts/Verdana-Italic.ttf");
                }
                else
                {
                    return new FontResolverInfo("Fonts/Verdana-Regular.ttf");
                }
            }
            return null;
        }
    }
    
    
    internal class MyImageSourceImp : ImageSource.IImageSource
    {
    public  MyImageSourceImp(string name)
     {
      Name = name;
     }
     public void SaveAsJpeg(MemoryStream ms)
     {
         string path =
             "/Users/dannybolick/MainNas/PersonalRepo/GeneralTestProjectsMain/PdfSharpCoreFork/PDFGen_test01/bin/Debug/output.png";
         File.WriteAllBytes(path,ms.ToArray());
         var test = "";
     }

     public bool Transparent { get; } = false;
     public void SaveAsPdfBitmap(MemoryStream ms)
     {
         var test = "";
     }

     public int Width { get; } = 100;
     public int Height { get; } = 100;
     public string Name { get; } = "test2.png";
    }
    
    internal class MyImageSource : ImageSource
    {
     protected override IImageSource FromFileImpl(string path, int? quality)
     {
       return new MyImageSourceImp(path);
     }
    
     protected override IImageSource FromBinaryImpl(string name, Func<byte[]> imageSource, int? quality)
     {
      return new MyImageSourceImp(name);
     }
    
     protected override IImageSource FromStreamImpl(string name, Func<Stream> imageStream, int? quality)
     {
      return new MyImageSourceImp(name);
     }
    }
}