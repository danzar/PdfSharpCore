using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Shapes;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using PdfSharpCore.Charting;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using FillFormat = PdfSharpCore.Charting.FillFormat;
using VerticalAlignment = PdfSharpCore.Charting.VerticalAlignment;


namespace CloudCare.BL.DataAccess.Services.ReportingService.Reports
{
  /// <summary>
  ///   <example>
  ///     Using MigraDoc (for fluent easiness) and PdfSharp (for complex graphics)
  ///     Documentation: http://www.pdfsharp.net/wiki/
  ///     Samples: http://www.pdfsharp.net/wiki/MixMigraDocAndPdfSharp-sample.ashx
  ///   </example>
  ///   <remarks>
  ///     This class is the base reusable class for all CloudCare generated reports.
  ///   </remarks>
  /// </summary>
  public class CloudCareReport
  {
    
    //CC-1734 set table records limit to avoid stack on big reports
    protected static int TableRowLimit = 500;
    protected static bool EnableTableLimit = true;

#region Prop/Fields

    /// <summary>
    ///   Containers for bar graphs
    /// </summary>
    // private readonly IList<BarGraphInfo> _barGraphs = new List<BarGraphInfo>();

    // private readonly IConfigurationService _configurationService;
    private readonly Document _document;
    // private readonly IList<MultiBarGraphInfo> _multiBarGraphs = new List<MultiBarGraphInfo>();

    private readonly Color _grayHeadlineColor = new Color(102, 102, 102);
    private readonly Color _datagridHeaderBackground = new Color(240, 240, 240);
    private readonly Color _filterBackground = new Color(245, 245, 245);
    private readonly Color _tableBorderColor = new Color(240, 240, 240);
    private readonly Color _oddRowColor = new Color(240, 240, 240);

    //size constants in milimeters
    private readonly int _headerWidth = 182;
    private readonly int _marginBetweenImageAndHeader = 5;
    private readonly int _defaultLogoWidth = 55;
    private readonly int _filterWidth = 202;
    private readonly int _filterKeysWidth = 40;

    private readonly int tableWidth = 202;

    private readonly float _defaultLogoAspectRatio = 4;

    /// <summary>
    ///   Container for pie charts
    /// </summary>
    //private readonly IList<PieChartInfo> _pieCharts = new List<PieChartInfo>();

    private readonly double offset = 12;

    //private LogoImageResource logoImageResouce = new LogoImageResource();
    private bool _isInitialized;
    private double headerAdjustment;

    public int ReportId { get; set; }

    public PdfDocument Document { get; set; }
    // publically exposed properties
    public string LogoUrl { get; set; }
    private string footImagePath { get; set; }

    public string DocumentPath { get; private set; }

    public bool ExtraPage { get; set; }

    // public Dictionary<ReportHeaderSectionItem, string> ReportHeaderItems { get; set; }
    // public Dictionary<ReportFilterSectionItem, string> ReportFilterSectionItems { get; set; }
    //
    // public Dictionary<ReportFooterSectionItem, string> ReportFooterSectionItems { get; set; }

    public  string customHeaderText { get; set; }

    #endregion



    public CloudCareReport()
    {
      _isInitialized = false;


      // this is the pdfDocument we will eventually output and send to our users
      _document = new Document();
      //The Header distance from the top
     // _document.DefaultPageSetup.HeaderDistance = Unit.FromMillimeter(5);


      //This allows us to have space from the header on the second page, but also bumps the space on the subject line
      // _document.DefaultPageSetup.TopMargin = Unit.FromMillimeter(25);
      //
      // _document.DefaultPageSetup.LeftMargin = Unit.FromMillimeter(5);//5  //Org 15
      // _document.DefaultPageSetup.RightMargin = Unit.FromMillimeter(5);//5 // org 15
      //
      // _document.DefaultPageSetup.FooterDistance = Unit.FromMillimeter(5);



      DefineStyles(); // defines all the default styles for the document (must be overriden when desired)
    }



    #region Initialize

    private void Initialize()
    {

      if (!_isInitialized)
      {
        _isInitialized = true;
        lock (_document)
        {
          var sec = _document.AddSection();

          WritePageHeader(sec);
          CompanyHeader(sec);
          WriteFooter(sec);
        }

        if (ExtraPage)
        {
          var sec2 = _document.AddSection();

          WritePageHeader(sec2);
          //ExtraPageSecondaryHeader(sec2);
          WriteFooter(sec2);
        }
      }
    }

    #endregion

    #region Create(string fileName) and CreateDocument(PdfDocument pdfDocument = null, string fileName = "")
    //Method called by all reports for creating the report.......
    public void Create(string fileName)
    {
      CreateDocument(fileName);
    }


    public void CreateDocument(string fileName)
    {
     
      Initialize(); //sets up the headers



      if (String.IsNullOrEmpty(fileName) || String.IsNullOrWhiteSpace(fileName))
      {
        throw new ArgumentNullException("fileName");
      }

      if (fileName.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) < 0)
      {
        fileName += ".pdf";
      }

     
      try
      {

     
        // Create a temporary file that will eventually be sent to the user  
        //Added "AddGuid" to the file name so that the file would not be used by multi processes
         DocumentPath = "SOME_PATH";

        PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);

        renderer.Document = _document.Clone();
        renderer.RenderDocument();
        // Save the document...

        renderer.PdfDocument.Save(DocumentPath);

        // right now, all graphs are placed on the first page.  later, this will be easy to change if we need to.
        #region Add Graphs to Doc
        // if (_barGraphs.Any() || _multiBarGraphs.Any() || _pieCharts.Any())
        // {
        //
        //   PdfDocument alteredPdf = PdfReader.Open(DocumentPath, PdfDocumentOpenMode.Modify);  
        //     foreach (var graph in _barGraphs)
        //     {
        //       graph.Location = new XPoint(graph.Location.X, graph.Location.Y + headerAdjustment);
        //
        //       DrawSimpleBarGraph(alteredPdf.Pages[0], graph);
        //     }
        //
        //     foreach (var multiBarGraph in _multiBarGraphs)
        //     {
        //       if (multiBarGraph.NewPage)
        //       {
        //         multiBarGraph.Location = new XPoint(multiBarGraph.Location.X, multiBarGraph.Location.Y);
        //       }
        //       else
        //       {
        //         multiBarGraph.Location = new XPoint(multiBarGraph.Location.X, multiBarGraph.Location.Y + headerAdjustment);
        //       }
        //
        //       // DrawMultiBarGraph(alteredPdf.Pages[0], multiBarGraph);
        //       DrawMultiBarGraph(alteredPdf, multiBarGraph);
        //     }
        //
        //     foreach (var pieChart in _pieCharts)
        //     {
        //       pieChart.Location = new XPoint(pieChart.Location.X, pieChart.Location.Y + headerAdjustment);
        //       DrawPieChart(alteredPdf.Pages[0], pieChart);
        //     }
        //   alteredPdf.Save(DocumentPath);
        // }

          #endregion
      }
      catch (Exception e)
      {
        throw;
      }

      // always gets deleted

 
    }

    #endregion



    #region Header/Footer section()

    private void DefineStyles()
    {
      // Get the predefined style Normal.
      var style = _document.Styles["Normal"];
      // Because all styles are derived from Normal, the next line changes the 
      // font of the whole document. Or, more exactly, it changes the font of
      // all styles and paragraphs that do not redefine the font.
      style.Font.Name = "Arial";
      style.Font.Size = Unit.FromPoint(8);
      style.Font.Color = new Color(0x50,0x50,0x50);//new Color(0x0, 0x34, 0x63);
      style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
      style.ParagraphFormat.LineSpacing = Unit.FromMillimeter(5);
      style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
    }

  

    private void WritePageHeader(Section sec)
    {
      sec.PageSetup.DifferentFirstPageHeaderFooter = false;

      var pageHeader = sec.Headers.Primary;

      Table secTable = pageHeader.AddTable();

      //default columns widths 
      Column columnMargin = secTable.AddColumn(Unit.FromMillimeter(2));
      Column columnImage = secTable.AddColumn();
      Column columnText = secTable.AddColumn();

      //Adding Image Area


      
      Row row = secTable.AddRow();
      
     // MigraImage cellImage = null;

      Cell cellItemImage = row.Cells[1];
      cellItemImage.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;
      Cell cellItemText = row.Cells[2];
      cellItemText.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;

       //Use custom Image
      {
        
       // cellImage = row.Cells[1].AddImage(logoImageResouce.imagePath);

        int height;
        int width;
        float aspect;

        // //get dimensions to determine size
        // using (var imageStream = File.OpenRead(logoImageResouce.imagePath))
        // {
        //   var decoder = BitmapDecoder.Create(imageStream, BitmapCreateOptions.IgnoreColorProfile,
        //     BitmapCacheOption.Default);
        //   height = decoder.Frames[0].PixelHeight;
        //   width = decoder.Frames[0].PixelWidth;
        //   aspect = (float)width / height;
        // }

        // if (aspect > _defaultLogoAspectRatio-0.1f)
        // {
        //   cellImage.Width = Unit.FromMillimeter(_defaultLogoWidth);
        //   cellImage.Height = cellImage.Width/aspect;
        // }
        // else
        // {
        //   cellImage.Height = Unit.FromMillimeter(_defaultLogoWidth/_defaultLogoAspectRatio);
        //   cellImage.Width = cellImage.Height * aspect;
        // }
      }

   //   columnImage.Width = Unit.FromMillimeter(cellImage.Width.Millimeter + _marginBetweenImageAndHeader);
      columnText.Width = Unit.FromMillimeter(_headerWidth - columnImage.Width.Millimeter);
      
      // //Add custom text if there is any
      // if (!customHeaderText.IsNullOrUndefined())
      // {
      //   //Custom Header Text 
      //   Cell headerTextColumn = row.Cells[2];
      //   headerTextColumn.Format.Font.Size = Unit.FromPoint(17);
      //
      //   Paragraph customTextParagraph = headerTextColumn.AddParagraph();
      //   customTextParagraph.Format.Borders.DistanceFromTop = Unit.FromPoint(4);
      //   customTextParagraph.AddFormattedText(customHeaderText);
      // }
        
      pageHeader.Format.Borders.DistanceFromTop = Unit.FromMillimeter(3);
      pageHeader.Format.Borders.DistanceFromLeft = Unit.FromMillimeter(5);
      

    }

    
    private void WriteFooter(Section sec)
    {
       sec.Footers.Primary.Add(CreateFooter());
    }


    private Table  CreateFooter()
    {
      var footerTable = new Table
      {
        Rows = { LeftIndent = 0 },
        Format = { Alignment = ParagraphAlignment.Center },
       // Shading = { Visible = true, Color = new Color(237, 237, 237) }
      };
      

        footerTable.TopPadding = Unit.FromMillimeter(0);
        footerTable.BottomPadding = Unit.FromMillimeter(0);

        footerTable.Format.Borders.DistanceFromTop = Unit.FromMillimeter(3);
        footerTable.Format.Borders.DistanceFromLeft = Unit.FromMillimeter(5);

        footerTable.Format.Font.Color = new Color(0, 0x34, 0x63);
        var footerHeight = Unit.FromPoint(20);
        var fontSize = Unit.FromPoint(8);
        var textBorderDistance = Unit.FromPoint(8);

        Column columnMargin = footerTable.AddColumn(Unit.FromPoint(0));

        Column columnImage = footerTable.AddColumn(Unit.FromPoint(120));
        columnImage.Format.Alignment = ParagraphAlignment.Left;
   
        Column columnTradeMark = footerTable.AddColumn(Unit.FromPoint(130));
        columnTradeMark.Format.Alignment = ParagraphAlignment.Left;
        Column columnPage = footerTable.AddColumn(Unit.FromPoint(130));
        columnPage.Format.Alignment = ParagraphAlignment.Center;
        Column columnGenerated = footerTable.AddColumn(Unit.FromPoint(170));
        columnGenerated.Format.Alignment = ParagraphAlignment.Right;


        Row row = footerTable.AddRow();

        // row.Borders.Color = Colors.Black;
        Cell marginCell = row.Cells[0];
        Cell imageCell = row.Cells[1];
        Cell tradeMarkCell = row.Cells[2];
        Cell pageTextCell = row.Cells[3];
        Cell generatedCell = row.Cells[4];


   
        //Image Area
    //    TextFrame imageTextFrame = imageCell.AddTextFrame();
    //    imageTextFrame.Height = footerHeight;
     //   imageTextFrame.MarginRight = Unit.FromMillimeter(10);
        string defaultUrl = ""; //this will make the url path return the default LOGO url.
        bool customImage;
        // footImagePath = defaultUrl.GetLogoImageFromHttpUrl(this.ReportId, out customImage);
        // //  MigraImage imageItem = imageTextFrame.AddImage(footImagePath);
        // MigraImage imageItem = imageCell.AddImage(footImagePath);
        // imageItem.LockAspectRatio = true;
        // imageItem.Height = Unit.FromPoint(20);
      //  imageItem.Width = Unit.FromPoint(110);





        #region TradeMark Area

        TextFrame tradeTextFrame = tradeMarkCell.AddTextFrame();
        tradeTextFrame.Height = footerHeight;
        tradeTextFrame.Width = Unit.FromPoint(130);
        Paragraph tradeParagraph = tradeTextFrame.AddParagraph();
        tradeParagraph.Format.Font.Size = fontSize;
        tradeParagraph.Format.Font.Name = "Arial";
      //  tradeParagraph.Format.Borders.DistanceFromTop = textBorderDistance;
        tradeParagraph.AddFormattedText("TRADEMARK TEXT");
        #endregion

        #region  Page# Area
        TextFrame pageTextFrame = pageTextCell.AddTextFrame();
        pageTextFrame.Height = footerHeight;
        Paragraph pageParagraph = pageTextFrame.AddParagraph();
        pageParagraph.Format.Font.Size = fontSize;
        pageParagraph.Format.Font.Name = "Arial";
   //     pageParagraph.Format.Borders.DistanceFromTop = textBorderDistance;
        pageParagraph.AddFormattedText("PAGE NUMBER:");
        pageParagraph.AddPageField();
        #endregion

        #region  Generated Area
        TextFrame generateTextFrame = generatedCell.AddTextFrame();
     //   generateTextFrame.FillFormat.Color = Colors.LightBlue;
        generateTextFrame.Height = footerHeight;
        generateTextFrame.Width = Unit.FromPoint(180);
        Paragraph generatedParagraph = generateTextFrame.AddParagraph();
        generatedParagraph.Format.Font.Size = fontSize;
        generatedParagraph.Format.Font.Name = "Arial";
    //    generatedParagraph.Format.Borders.DistanceFromTop = textBorderDistance;
        generatedParagraph.Format.Alignment = ParagraphAlignment.Right;
        generatedParagraph.AddFormattedText("GENERATE DAT$E TEXT");
        #endregion




      
      return footerTable;
    }
    #endregion

    #region Write Header/Filter Text
    private void CompanyHeader(Section sec)
    {
      //Title
      Paragraph titleHeader = sec.AddParagraph();
      titleHeader.Format.LeftIndent = Unit.FromMillimeter(-1);
      FormattedText formattedText = titleHeader.AddFormattedText("TITLE", TextFormat.Bold);
      formattedText.Font.Size = 16;
      formattedText.Color = _grayHeadlineColor;

      //Subject
      // if (ReportHeaderItems.ContainsKey(ReportHeaderSectionItem.Subject) && !ReportHeaderItems[ReportHeaderSectionItem.Subject].IsNullOrUndefined())
      // {
        AddVerticalSpace(sec,Unit.FromMillimeter(2));
        var subjectHeader = sec.AddParagraph();
        subjectHeader.Format.LeftIndent = Unit.FromMillimeter(-1); // this is needed to align 
        FormattedText subjectText = subjectHeader.AddFormattedText("SUBJECT");
        subjectText.Font.Size = 11;
        subjectText.Color = _grayHeadlineColor;
      // }
      // else
      // {
      //   AddLineBreaks(sec.AddParagraph());
      // }
      
      //Filter


 
    }



    private void WriteFilterRow(Table table,string key,string value)
    {
      Row row = table.AddRow();
      row.BottomPadding = Unit.FromMillimeter(1);
      row.TopPadding = Unit.FromMillimeter(1);
      //row.Shading.Color = _filterBackground;

      var cellText = row.Cells[0].AddParagraph();
      
      cellText.Format.LeftIndent = Unit.FromPoint(3);
      FormattedText cellTextFormat = cellText.AddFormattedText(key, TextFormat.Bold);
      cellTextFormat.Font.Size = 12;
      cellText.Format.Font.Color = _grayHeadlineColor;
      cellText.Format.Alignment = ParagraphAlignment.Left;
      
      var cellValue = row.Cells[1].AddParagraph();
      cellValue.Format.Alignment = ParagraphAlignment.Left;
      cellValue.Format.Font.Color = _grayHeadlineColor;
      cellValue.Format.Borders.DistanceFromLeft = Unit.Zero;
      FormattedText cellValueFormattedText = cellValue.AddFormattedText(value);
      cellValueFormattedText.Font.Size = 12;
    }


    #endregion

    #region AddDataGrid<T>(IEnumerable<T> rows, string[] headerFields, int[] columnWidths)

    public void AddDataGrid<T>(IEnumerable<T> rows, string[] headerFields, double[] columnWidths)
    {
      PrintRows(rows, headerFields, columnWidths);
    }

    #region PrintHeader(string[] headerFields, double[] columnWidths)

    private void PrintHeader(string[] headerFields, double[] columnWidths, Table table)
    {
      for (var i = 0; i < headerFields.Length; i++)
      {
        var column = table.AddColumn(columnWidths[i].ToString() + "mm");
        column.Format.Alignment = ParagraphAlignment.Left;
        column.Format.LeftIndent = Unit.FromMillimeter(0.5);
        column.Format.RightIndent = Unit.FromMillimeter(0.5);
      }

      var row = table.AddRow();
      row.HeadingFormat = true;
      row.HeightRule = RowHeightRule.AtLeast;
      row.Height = Unit.FromMillimeter(10);
      row.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;
      row.Format.Shading.Color =  _datagridHeaderBackground;// new Color(0xf0, 0xf0, 0xf0);//Inner Text area
      row.Format.Borders.Bottom.Width = 0;
      row.HeadingFormat = true;
      row.Shading.Color = _datagridHeaderBackground;//  new Color(0xf0, 0xf0, 0xf0);//Outer Text area

      

      table.LeftPadding = Unit.FromMillimeter(1);
      table.TopPadding = Unit.FromMillimeter(1);
      table.BottomPadding = Unit.FromMillimeter(1);

      //
     // table.TopPadding = Unit.FromMillimeter(55);

      for (var i = 0; i < table.Columns.Count; i++)
      {
        if (i == 0)
        {
          row.Cells[i].Borders.Left.Color = _datagridHeaderBackground;
        }
        if (i == table.Columns.Count - 1)
        {
          row.Cells[i].Borders.Right.Color = _datagridHeaderBackground;
        }
        //row.Cells[i].Format.Shading.Color = tableBorderColor;
        //row.Cells[i].Format.Borders.Bottom.Color = new Color(0xe2, 0xe1, 0xe1);
        var para = row.Cells[i].AddParagraph(headerFields[i]);
        para.Format.Font.Name = "Arial";
        para.Format.Font.Size = Unit.FromPoint(10);
        para.Format.LineSpacingRule = LineSpacingRule.AtLeast;
        para.Format.LineSpacing = Unit.FromMillimeter(3);

        //para.Format.Borders.Bottom.Color = new Color(0xe2, 0xe1, 0xe1);
      }

      var gapRow = table.AddRow();
      gapRow.HeadingFormat = true;
      gapRow.HeightRule = RowHeightRule.Exactly;
      gapRow.Height = Unit.FromMillimeter(1);
      gapRow.HeadingFormat = true;

      gapRow.Cells[0].Borders.Left.Width = Unit.FromMillimeter(.5);
      gapRow.Cells[0].Borders.Left.Color = _tableBorderColor;

      gapRow.Shading.Color = new Color(0xe1, 0xe1, 0xe1);

      gapRow.Cells[table.Columns.Count - 1].Borders.Right.Width = Unit.FromMillimeter(.5);
      gapRow.Cells[table.Columns.Count - 1].Borders.Right.Color = _tableBorderColor;
    }

    #endregion

    private void PrintRows<T>(IEnumerable<T> rows, string[] headerFields, double[] cellWidths)
    {
      Initialize(); //sets up the headers
      
      var sec = _document.LastSection;

      // if (position == DataGridPositions.ReportMiddle)
      // {
      //   AddLineBreaks(sec.LastParagraph, CloudCarePdfStyles.HalfPageDataGridSpace);
      // }
      //
      // if (position == DataGridPositions.ReportNextPage)
      // {
      //   sec.AddPageBreak();
      // }

      var genericInstance = Activator.CreateInstance<T>();
      var properties = genericInstance.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);

      T[] dataRows = new T[0];

      if (rows != null) dataRows = rows as T[] ?? rows.ToArray();
      int rowsCount = dataRows.Length;

      #region CC-1734 table_size_limit

      if (EnableTableLimit)
      {
        //set table records limit to avoid stack on big reports
        rowsCount = Math.Min(dataRows.Length, TableRowLimit);

        //add note to the document if table was truncated
        if (rowsCount != dataRows.Length)
        {
          
        }
      }

      #endregion CC-1734 table_size_limit
      
      var table = sec.AddTable();
      table.Style = "Table";
      table.Borders.Width = Unit.FromMillimeter(.5);
      table.Borders.Color = new Color(0xff, 0xff, 0xff);
      table.Borders.Bottom.Width = Unit.FromMillimeter(0);
      table.Borders.Top.Width = Unit.FromMillimeter(0);
      table.LeftPadding = 0;
      table.RightPadding = 0;
      table.TopPadding = 0;
      table.BottomPadding = 0;

      PrintHeader(headerFields, cellWidths, table);

      for (var i = 0; i < rowsCount; i++)
      {
        var rowInstance = table.AddRow();

        if (i%2 != 0)
        {
          rowInstance.Shading.Color = _oddRowColor;
          rowInstance.Borders.Color = _oddRowColor;
        }
        rowInstance.HeadingFormat = false;
        rowInstance.Height = Unit.FromMillimeter(6);
        rowInstance.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;
        rowInstance.Borders.Bottom.Width = Unit.FromMillimeter(.5);
        rowInstance.Borders.Bottom.Color = _oddRowColor;

        for (var j = 0; j < properties.Length; j++)
        {
          /*if (i != 0)
          {
            rowInstance.Cells[j].Borders.Top.Color = new Color(0xf0, 0xf0, 0xf0);
          }*/
          if (j == 0)
          {
            rowInstance.Cells[j].Borders.Left.Color = _tableBorderColor;
          }
          if (j == table.Columns.Count - 1)
          {
            rowInstance.Cells[j].Borders.Right.Color = _tableBorderColor;
          }

          if (i + 1 == dataRows.Count())
          {
            rowInstance.Cells[j].Borders.Bottom.Color = _tableBorderColor;
          }

          var propName = properties.ElementAt(j).Name;
          var cellValue =
            dataRows.ElementAt(i)
              .GetType()
              .GetProperty(propName)
              .GetValue(dataRows.ElementAt(i), null);
          if (cellValue != null)
          {
            try
            {
              Cell currentCell = rowInstance.Cells[j];

              // if (customCellRenderer != null)
              // {
              //   customCellRenderer(currentCell, propName, cellValue); 
              // } else
              // {
                DefaultCellRenderer(currentCell, propName, cellValue);
              //}
              
            }
            catch (Exception e)
            {
              throw new Exception($"Index Item:{j} Error:{e} Inner:{e.InnerException} Stack:{e.StackTrace}");
            }
          }
        }
      }

    }

    #endregion


    /// <summary>
    /// Add the text to the cell with spaces(line breaks) where needed. 
    /// </summary>
    /// <param name="instring"></param>
    /// <param name="cell"></param>
    /// <param name="fontName"></param>
    /// <param name="fontsize"></param>
    /// <returns></returns>
    private static string AddTextToCell(string instring, Cell cell,string fontName, Unit fontsize)
    {
      PdfDocument pdfd = new PdfDocument();
      PdfPage pg = pdfd.AddPage();
      XGraphics oGFX = XGraphics.FromPdfPage(pg);
      Unit maxWidth = cell.Column.Width - (cell.Column.LeftPadding + cell.Column.RightPadding);
      Paragraph par;

      //Added to prevent font size of 0 or Max value being passed to new font.
      if (fontsize.IsEmpty || fontsize.Value <= 0 || fontsize.Value > Single.MaxValue)
      {
        fontsize = Unit.FromPoint(10);
      }
      XFont font = new XFont(fontName, fontsize);
     

      //if string is to long or not
      if (oGFX.MeasureString(instring, font).Width > maxWidth.Value)
      {
        int cellMax = cellMaxChars(cell, font);
        
        // the string has no spaces in it.
        if (instring.IndexOf(' ') == -1)
        {
          var splitSting = splitStringToCell(instring, cell, font);

          //Insert spaces as joiners for the list of strings.
          instring = string.Join(" ", splitSting);      
        }
        else //Deal with strings that has spaces in them also
        {
          string outputSting = "";

          //where we are checking to know if we want to break at the space.
          double breakPoint = Math.Ceiling(cellMax * .75);

          //split string up for the cell
          var splitSting = splitStringToCell(instring, cell, font);

          //Join string back together with the spaces needed.
          foreach (var item in splitSting)
          {
            //If there is no space in string OR space is less than break point insert space Else let the let the existing space break the line
            if (item.IndexOf(' ') == -1 || item.LastIndexOf(' ') < breakPoint)
            {
              outputSting += item + " ";
            }
            else
            {
              outputSting += item;
            }
          }


          instring = outputSting;
        }

      }
   
      return instring;
    
    }

    /// <summary>
    /// Split the string up to match the cell width
    /// </summary>
    /// <param name="input"></param>
    /// <param name="cell"></param>
    /// <param name="font"></param>
    /// <returns></returns>
    private static List<string> splitStringToCell(string input, Cell cell, XFont font)
    {

      PdfDocument pdfd = new PdfDocument();
      PdfPage pg = pdfd.AddPage();
      XGraphics oGFX = XGraphics.FromPdfPage(pg);
      Unit maxWidth = cell.Column.Width - (cell.Column.LeftPadding + cell.Column.RightPadding);

      //This is to make the space at the need for the needed line breaks
      XSize spaceSize = oGFX.MeasureString("XX",font);

      List<string> returnList = new List<string>();
      char[] inputCharArray = input.ToCharArray();

      string tempString = "";

      int Multiplier = 1;

      for (int i = 0; i < inputCharArray.Length; i++)
      {
         char currentChar = inputCharArray[i];

        if (oGFX.MeasureString(tempString + currentChar, font).Width < (maxWidth.Value - spaceSize.Width))
        {
          tempString += inputCharArray[i];

          if (i == inputCharArray.Length - 1)
          {
            returnList.Add(tempString);
          }

        }
        else
        {
          returnList.Add(tempString);
          tempString = inputCharArray[i].ToString();
        }
      }

      return returnList;
      
    }

    /// <summary>
    /// Get the Cell max char width
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="font"></param>
    /// <returns></returns>
    private static int cellMaxChars(Cell cell, XFont font)
    {
      PdfDocument pdfd = new PdfDocument();
      PdfPage pg = pdfd.AddPage();
      XGraphics oGFX = XGraphics.FromPdfPage(pg);
      Unit maxWidth = cell.Column.Width - (cell.Column.LeftPadding + cell.Column.RightPadding);

      int returnInt = 0;

      string testString = "";

      while (oGFX.MeasureString(testString,font).Width < maxWidth.Value)
      {
        testString += "X";
        returnInt++;
      }


      return returnInt;

    }

    public void AddHighlightsBar(string header, IEnumerable<Highlight> highlights)
    {
      Initialize(); //sets up the headers

      if (highlights == null)
      {
        return;
      }

      var sec = _document.LastSection;
      
      var table = sec.AddTable();
      table.Style = "Table";
      /*table.LeftPadding = 0;
      table.RightPadding = 0;
      table.TopPadding = 0;
      table.BottomPadding = 0;*/

      var highlightsArray = highlights as Highlight[] ?? highlights.ToArray();
      int count = highlightsArray.Length;
      var singleWidth = Unit.FromMillimeter((float)tableWidth / count);

      for (var i = 0; i < count; i++)
      {
        var column = table.AddColumn(singleWidth);
        column.Format.Alignment = ParagraphAlignment.Center;
        column.Format.LeftIndent = Unit.FromMillimeter(1);
        column.Format.RightIndent = Unit.FromMillimeter(1);
      }

      var headingRow = table.AddRow();
      headingRow.HeadingFormat = true;
      headingRow.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;//VerticalAlignment.Center;
      headingRow.Shading.Color = new Color(0xe1, 0xe1, 0xe1);// new Color(0xf0, 0xf0, 0xf0);//Inner Text area

      //make shared heading
      headingRow.Cells[0].MergeRight = count - 1;
      var para = headingRow.Cells[0].AddParagraph(header);
      para.Format.Font.Name = "Arial";
      para.Format.Alignment = ParagraphAlignment.Left;
      para.Format.Font.Size = Unit.FromPoint(10);
      para.Format.SpaceBefore = Unit.FromMillimeter(2);
      para.Format.SpaceAfter = Unit.FromMillimeter(3);


      // Print highlights
      var rowInstance = table.AddRow();
    
      rowInstance.HeadingFormat = false;
      rowInstance.Height = Unit.FromMillimeter(6);
      rowInstance.VerticalAlignment = MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center;
      rowInstance.Shading.Color = _filterBackground;

      for (int i = 0; i < count; i++)
      {
        try
        {
          Cell currentCell = rowInstance.Cells[i];
          
          //value
          Paragraph cellValue = currentCell.AddParagraph(highlightsArray[i].Value);
          cellValue.Format.Font.Size = 20;
          cellValue.Format.Font.Bold = true;
          cellValue.Format.Font.Color = highlightsArray[i].Color;
          cellValue.Format.SpaceBefore = Unit.FromMillimeter(8);

          //description
          Paragraph cellDescription = currentCell.AddParagraph(highlightsArray[i].Description);
          cellDescription.Format.Font.Size = 12;
          cellDescription.Format.SpaceAfter = Unit.FromMillimeter(8);
        }
        catch (Exception e)
        {
          throw new Exception($"Index Item:{i} Error:{e} Inner:{e.InnerException} Stack:{e.StackTrace}");
        }
      }
      
      AddVerticalSpace(sec,Unit.FromMillimeter(5));
    }

    public void AddHeading(string text, double spaceBeforeInMillimeters = 3, double spaceAfterInMillimeters = 2)
    {
      var headingParagraph = _document.LastSection.AddParagraph();
      headingParagraph.Format.SpaceBefore = Unit.FromMillimeter(spaceBeforeInMillimeters);
      headingParagraph.Format.SpaceAfter = Unit.FromMillimeter(spaceAfterInMillimeters);
      headingParagraph.Format.LeftIndent = Unit.FromMillimeter(-1);
      ;
      FormattedText headingText = headingParagraph.AddFormattedText(text);

      headingText.Font.Size = 11;
      headingText.Font.Bold = true;
    }

    
    public void AddPageText(string text, double spaceBeforeInMillimeters = 3, double spaceAfterInMillimeters = 2)
    {
      Initialize(); //sets up the headers
      
      Paragraph pageTextPara = _document.LastSection.AddParagraph();
      pageTextPara.Format.SpaceBefore = Unit.FromMillimeter(spaceBeforeInMillimeters);
      pageTextPara.Format.SpaceAfter = Unit.FromMillimeter(spaceAfterInMillimeters);
      pageTextPara.Format.LeftIndent = Unit.FromMillimeter(-1);
      ;
      FormattedText headingText = pageTextPara.AddFormattedText(text);
      
      headingText.Font.Size = 11;
      headingText.Color = _grayHeadlineColor;
    }

    public static Paragraph DefaultCellRenderer(Cell currentCell, string propName, object value)
    {
      Paragraph cellParagraph = currentCell.AddParagraph(AddTextToCell(value.ToString(), currentCell, "Arial", 10));
      cellParagraph.Format.Font.Size = 10;

      var filler = currentCell.AddParagraph(" ");
      filler.Format.LineSpacing = Unit.FromMillimeter(1);

      return cellParagraph;
    }

    #region Add Graph to Doc Object
    // public void AddBarGraph(string header, Color[] colors, double[] values, string[] names, BarGraphStyles style)
    // {
    //   throw new NotImplementedException();
    // }



    public void AddVerticalSpace(Section section, Unit size)
    {
      Initialize();
      Paragraph paragraph = section.AddParagraph();
      paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
      paragraph.Format.LineSpacing = size;
    }

    public void AddVerticalSpace(Unit size)
    {
      Initialize();
      var section = _document.LastSection;
      AddVerticalSpace(section,size);
      
    }

    public Section GetLastSection()
    {
      return _document.LastSection;
    }

    #endregion

    private void AddLineBreaks(Paragraph paragraph, int amount = 1)
    {
      for (var i = 0; i < amount; i++)
      {
        paragraph.AddLineBreak();
      }
    }




    private void AddEmptyRow(Table table, Unit height)
    {
      Row row = table.AddRow();
      row.TopPadding = 0;
      row.BottomPadding = height;
      row.Format.LineSpacingRule = LineSpacingRule.Exactly;
      row.Format.LineSpacing = 0;
    }

    private void AddLineDivider(Section sec, double width = 565, double height = 0.5)
    {  //width was 505 with left/right margin at 15
     
      var lineDivider = sec.AddTextFrame();
      lineDivider.RelativeVertical = RelativeVertical.Line;
      lineDivider.MarginTop = Unit.FromPoint(0);
      lineDivider.MarginBottom = Unit.FromPoint(0);
      lineDivider.Width = width;
      lineDivider.Height = height;
      lineDivider.FillFormat = new MigraDocCore.DocumentObjectModel.Shapes.FillFormat();//new FillFormat {Color = XColors.Blue};
    }

    private void addTopSpacer(Section sec, int spaceAmount ,int width = 565)
    {
      var topper = sec.AddTextFrame();
      topper.MarginTop = Unit.FromPoint(0);
      topper.MarginBottom = Unit.FromPoint(0);
      topper.Width = width;
      topper.Height = spaceAmount;
      topper.RelativeVertical = RelativeVertical.Line;
    }

    private void addBottomSpacer(Section sec, int spaceAmount, int width = 565)
    {
      var bottom = sec.AddTextFrame();
      bottom.MarginTop = Unit.FromPoint(0);
      bottom.MarginBottom = Unit.FromPoint(0);
      bottom.Width = width;
      bottom.Height = spaceAmount;
      bottom.RelativeVertical = RelativeVertical.Line;
    }

    private void headerAdjustmentOffset(double amount)
    {
      headerAdjustment += amount;
    }
    
    // public void AddPieChart(PieChartInfo pieChartInfo)
    // {
    //   _pieCharts.Add(pieChartInfo);
    // }
    //
    // public void AddBarGraph(BarGraphInfo barGraphInfo)
    // {
    //   int hightest = 0;
    //   foreach (var barGraphDatum in barGraphInfo.BarGraphData)
    //   {
    //     if (barGraphDatum.Value > hightest)
    //     {
    //       hightest = (int)barGraphDatum.Value;
    //     }
    //   }
    //
    //   if (hightest > 0 && hightest < 5)
    //   {
    //     barGraphInfo.TickStep = 1;
    //   }
    //
    //
    //   _barGraphs.Add(barGraphInfo);
    // }
    //
    // public void AddMultiBarGraph(MultiBarGraphInfo multiBarGraphInfo)
    // {
    //   int hightest = 0;
    //   foreach (var multiBarGraphDatum in multiBarGraphInfo.MultiBarGraphData)
    //   {
    //     for (int i = 0; i < multiBarGraphDatum.Values.Length; i++)
    //     {
    //       if (multiBarGraphDatum.Values[i] > hightest)
    //       {
    //         hightest = (int)multiBarGraphDatum.Values[i];
    //       }
    //     }
    //   }
    //
    //   if (hightest > 0 && hightest < 5)
    //   {
    //     multiBarGraphInfo.TickStep = 1;
    //   }
    //
    //   _multiBarGraphs.Add(multiBarGraphInfo);
    // }
    //
    //
    //     #region drawing area
    //
    // private void DrawSimpleBarGraph(PdfPage page, BarGraphInfo barGraphInfo)
    // {
    //   using (var gfx = XGraphics.FromPdfPage(page))
    //   {
    //     // setup chart and frame
    //     var cf = new ChartFrame();
    //     var chart = new Chart(ChartType.Column2D);
    //
    //     // setup size related info
    //     if (barGraphInfo.Size == BarGraphSizes.Halfpage)
    //     {
    //       gfx.SetupHalfPageBarGraph(chart, cf, barGraphInfo);
    //     }
    //
    //     if (barGraphInfo.Size == BarGraphSizes.QuarterPage)
    //     {
    //       gfx.SetupQuarterPageBarGraph(chart, cf, barGraphInfo);
    //     }
    //
    //     // setup common graph styles
    //     chart.SetCommonGraphStyles(barGraphInfo);
    //
    //     gfx.DrawBarChart(chart, cf, barGraphInfo);
    //
    //     // render the graph object
    //     gfx.RenderObject();
    //   }
    // }
    //
    // private void DrawMultiBarGraph(PdfDocument doc, MultiBarGraphInfo multiBarGraphInfo)
    // {
    //   var page = doc.Pages[0];
    //   if (multiBarGraphInfo.NewPage)
    //   {
    //     if (doc.Pages.Count == 1)
    //     {
    //        page = doc.AddPage();
    //     }
    //     else
    //     {
    //       page = doc.Pages[1];
    //     }
    //   }
    //
    //
    //   using (var gfx = XGraphics.FromPdfPage(page))
    //   {
    //     var cf = new ChartFrame();
    //     var chart = new Chart(ChartType.Column2D);
    //
    //
    //     if (multiBarGraphInfo.Size == BarGraphSizes.Halfpage)
    //     {
    //       gfx.SetupHalfPageMultiBarGraph(chart, cf, multiBarGraphInfo);
    //     }
    //     if (multiBarGraphInfo.Size == BarGraphSizes.QuarterPage)
    //     {
    //       gfx.SetupQuarterPageMultiBarGraph(chart, cf, multiBarGraphInfo);
    //     }
    //     
    //    
    //     gfx.DrawMultiGraphLegend(chart, cf, multiBarGraphInfo);
    //
    //     gfx.RenderObject();
    //   }
    // }
    //
    //
    // private void DrawPieChart(PdfPage page, PieChartInfo pieChartInfo)
    // {
    //   using (var gfx = XGraphics.FromPdfPage(page))
    //   {
    //     var cf = new ChartFrame();
    //     var chart = new Chart(ChartType.Pie2D);
    //
    //     if (pieChartInfo.Size == PieChartSizes.HalfPage)
    //     {
    //       gfx.SetupHalfPagePieChart(chart, cf, pieChartInfo);
    //     }
    //     if (pieChartInfo.Size == PieChartSizes.QuarterPage)
    //     {
    //       gfx.SetupQuarterPagePieChart(chart, cf, pieChartInfo);
    //     }
    //
    //     gfx.DrawWhitePieChartRightLegend(chart, cf, pieChartInfo);
    //     gfx.RenderObject();
    //   }
    // }
    //
    // #endregion

  }
  
  public class Highlight
  {
    public string Value;
    public string Description;
    public Color Color;
  }


}