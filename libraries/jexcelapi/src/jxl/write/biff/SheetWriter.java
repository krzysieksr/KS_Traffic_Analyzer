/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
*
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
* Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl.write.biff;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.TreeSet;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.CellFeatures;
import jxl.CellReferenceHelper;
import jxl.Range;
import jxl.SheetSettings;
import jxl.WorkbookSettings;
import jxl.biff.AutoFilter;
import jxl.biff.ConditionalFormat;
import jxl.biff.DataValidation;
import jxl.biff.DataValiditySettingsRecord;
import jxl.biff.DVParser;
import jxl.biff.WorkspaceInformationRecord;
import jxl.biff.XFRecord;
import jxl.biff.drawing.Chart;
import jxl.biff.drawing.SheetDrawingWriter;
import jxl.biff.formula.FormulaException;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Blank;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableHyperlink;
import jxl.write.WriteException;

/**
 * Contains the functionality necessary for writing out a sheet.  Originally
 * this was incorporated in WritableSheetImpl, but was moved out into
 * a dedicated class in order to reduce the over bloated nature of that
 * class
 */
final class SheetWriter
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(SheetWriter.class);
    
  /**
   * A handle to the output file which the binary data is written to
   */
  private File outputFile;

  /**
   * The rows within this sheet
   */
  private RowRecord[] rows;

  /**
   * A number of rows.  This is a count of the maximum row number + 1
   */
  private int numRows;

  /**
   * The number of columns.  This is a count of the maximum column number + 1
   */
  private int numCols;

  /**
   * The page header
   */
  private HeaderRecord header;
  /**
   * The page footer
   */
  private FooterRecord footer;
  /**
   * The settings for the sheet
   */
  private SheetSettings settings;
  /**
   * The settings for the workbook
   */
  private WorkbookSettings workbookSettings;
  /**
   * Array of row page breaks
   */
  private ArrayList rowBreaks;
  /**
   * Array of column page breaks
   */
  private ArrayList columnBreaks;
  /**
   * Array of hyperlinks
   */
  private ArrayList hyperlinks;
  /**
   * Array of conditional formats
   */
  private ArrayList conditionalFormats;
  /**
   * The autofilter info
   */
  private AutoFilter autoFilter;
  /**
   * Array of validated cells
   */
  private ArrayList validatedCells;
  /**
   * The data validation validations
   */
  private DataValidation dataValidation;

  /**
   * The list of merged ranges
   */
  private MergedCells mergedCells;

  /**
   * The environment specific print record
   */
  private PLSRecord plsRecord;

  /**
   * The button property ste
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * The workspace options
   */
  private WorkspaceInformationRecord workspaceOptions;
  /** 
   * The column format overrides
   */
  private TreeSet columnFormats;

  /**
   * The list of drawings
   */
  private SheetDrawingWriter drawingWriter;

  /**
   * Flag indicates that this sheet contains just a chart, and nothing
   * else
   */
  private boolean chartOnly;

  /**
   * The maximum row outline level
   */
  private int maxRowOutlineLevel;

  /**
   * The maximum column outline level
   */
  private int maxColumnOutlineLevel;

  /**
   * A handle back to the writable sheet, in order for this class
   * to invoke the get accessor methods
   */
  private WritableSheetImpl sheet;


  /**
   * Creates a new <code>SheetWriter</code> instance.
   *
   * @param of the output file
   */
  public SheetWriter(File of,
                     WritableSheetImpl wsi,
                     WorkbookSettings ws)
  {
    outputFile = of;
    sheet = wsi;
    workspaceOptions = new WorkspaceInformationRecord();
    workbookSettings = ws;
    chartOnly = false;
    drawingWriter = new SheetDrawingWriter(ws);
  }

  /**
   * Writes out this sheet.  First writes out the standard sheet
   * information then writes out each row in turn.
   * Once all the rows have been written out, it retrospectively adjusts
   * the offset references in the file
   * 
   * @exception IOException 
   */
  public void write() throws IOException
  {
    Assert.verify(rows != null);

    // This worksheet consists of just one chart, so write it and return
    if (chartOnly)
    {
      drawingWriter.write(outputFile);
      return;
    }

    BOFRecord bof = new BOFRecord(BOFRecord.sheet);
    outputFile.write(bof);

    // Compute the number of blocks of 32 rows that will be needed
    int numBlocks = numRows / 32;
    if (numRows - numBlocks * 32 != 0)
    {
      numBlocks++;
    }

    int indexPos = outputFile.getPos();
   
    // Write the index record out now in order to serve as a place holder
    // The bof passed in is the bof of the workbook, not this sheet
    IndexRecord indexRecord = new IndexRecord(0, numRows, numBlocks);
    outputFile.write(indexRecord);

    if (settings.getAutomaticFormulaCalculation())
    {
      CalcModeRecord cmr = new CalcModeRecord(CalcModeRecord.automatic);
      outputFile.write(cmr);
    }
    else
    {
      CalcModeRecord cmr = new CalcModeRecord(CalcModeRecord.manual);
      outputFile.write(cmr);
    }

    CalcCountRecord ccr = new CalcCountRecord(0x64);
    outputFile.write(ccr);

    RefModeRecord rmr = new RefModeRecord();
    outputFile.write(rmr);
    
    IterationRecord itr = new IterationRecord(false);
    outputFile.write(itr);

    DeltaRecord dtr = new DeltaRecord(0.001);
    outputFile.write(dtr);

    SaveRecalcRecord srr = new SaveRecalcRecord
      (settings.getRecalculateFormulasBeforeSave());
    outputFile.write(srr);  

    PrintHeadersRecord phr = new PrintHeadersRecord
      (settings.getPrintHeaders());
    outputFile.write(phr);

    PrintGridLinesRecord pglr = new PrintGridLinesRecord
      (settings.getPrintGridLines());
    outputFile.write(pglr);

    GridSetRecord gsr = new GridSetRecord(true);
    outputFile.write(gsr);

    GuttersRecord gutr = new GuttersRecord();
    gutr.setMaxColumnOutline(maxColumnOutlineLevel + 1);
    gutr.setMaxRowOutline(maxRowOutlineLevel + 1);

    outputFile.write(gutr);

    DefaultRowHeightRecord drhr = new DefaultRowHeightRecord
     (settings.getDefaultRowHeight(), 
      settings.getDefaultRowHeight() != 
                SheetSettings.DEFAULT_DEFAULT_ROW_HEIGHT);
    outputFile.write(drhr);

    if (maxRowOutlineLevel > 0)
    {
      workspaceOptions.setRowOutlines(true);
    }

    if (maxColumnOutlineLevel > 0)
    {
      workspaceOptions.setColumnOutlines(true);
    }

    workspaceOptions.setFitToPages(settings.getFitToPages());
    outputFile.write(workspaceOptions);

    if (rowBreaks.size() > 0)
    {
      int[] rb = new int[rowBreaks.size()];

      for (int i = 0; i < rb.length; i++)
      {
        rb[i] = ( (Integer) rowBreaks.get(i)).intValue();
      }

      HorizontalPageBreaksRecord hpbr = new HorizontalPageBreaksRecord(rb);
      outputFile.write(hpbr);
    }

    if (columnBreaks.size() > 0)
    {
      int[] rb = new int[columnBreaks.size()];

      for (int i = 0; i < rb.length; i++)
      {
        rb[i] = ( (Integer) columnBreaks.get(i)).intValue();
      }

      VerticalPageBreaksRecord hpbr = new VerticalPageBreaksRecord(rb);
      outputFile.write(hpbr);
    }

    HeaderRecord header = new HeaderRecord(settings.getHeader().toString());
    outputFile.write(header);

    FooterRecord footer = new FooterRecord(settings.getFooter().toString());
    outputFile.write(footer);

    HorizontalCentreRecord hcr = new HorizontalCentreRecord
      (settings.isHorizontalCentre());
    outputFile.write(hcr);

    VerticalCentreRecord vcr = new VerticalCentreRecord
      (settings.isVerticalCentre());
    outputFile.write(vcr);

    // Write out the margins if they don't equal the default
    if (settings.getLeftMargin() != settings.getDefaultWidthMargin())
    {
      MarginRecord mr = new LeftMarginRecord(settings.getLeftMargin());
      outputFile.write(mr);
    }

    if (settings.getRightMargin() != settings.getDefaultWidthMargin())
    {
      MarginRecord mr = new RightMarginRecord(settings.getRightMargin());
      outputFile.write(mr);
    }

    if (settings.getTopMargin() != settings.getDefaultHeightMargin())
    {
      MarginRecord mr = new TopMarginRecord(settings.getTopMargin());
      outputFile.write(mr);
    }

    if (settings.getBottomMargin() != settings.getDefaultHeightMargin())
    {
      MarginRecord mr = new BottomMarginRecord(settings.getBottomMargin());
      outputFile.write(mr);
    }

    if (plsRecord != null)
    {
      outputFile.write(plsRecord);
    }

    SetupRecord setup = new SetupRecord(settings);
    outputFile.write(setup);

    if (settings.isProtected())
    {
      ProtectRecord pr = new ProtectRecord(settings.isProtected());
      outputFile.write(pr);

      ScenarioProtectRecord spr = new ScenarioProtectRecord
        (settings.isProtected());
      outputFile.write(spr);

      ObjectProtectRecord opr = new ObjectProtectRecord
        (settings.isProtected());
      outputFile.write(opr);

      if (settings.getPassword() != null)
      {
        PasswordRecord pw = new PasswordRecord(settings.getPassword());
        outputFile.write(pw);
      }
      else if (settings.getPasswordHash() != 0)
      {
        PasswordRecord pw = new PasswordRecord(settings.getPasswordHash());
        outputFile.write(pw);
      }
    }

    indexRecord.setDataStartPosition(outputFile.getPos());
    DefaultColumnWidth dcw = 
      new DefaultColumnWidth(settings.getDefaultColumnWidth());
    outputFile.write(dcw);
    
    // Get a handle to the normal styles
    WritableCellFormat normalStyle = 
      sheet.getWorkbook().getStyles().getNormalStyle();
    WritableCellFormat defaultDateFormat = 
      sheet.getWorkbook().getStyles().getDefaultDateFormat();

    // Write out all the column formats
    ColumnInfoRecord cir = null;
    for (Iterator colit = columnFormats.iterator(); colit.hasNext() ; )
    {
      cir = (ColumnInfoRecord) colit.next();

      // Writing out the column info with index 0x100 causes excel to crash
      if (cir.getColumn() < 0x100)
      {
        outputFile.write(cir);
      }

      XFRecord xfr = cir.getCellFormat();
      
      if (xfr != normalStyle && cir.getColumn() < 0x100)
      {
        // Make this the format for every cell in the column
        Cell[] cells = getColumn(cir.getColumn());

        for (int i = 0; i < cells.length; i++)
        {
          if (cells[i] != null &&
              (cells[i].getCellFormat() == normalStyle ||
               cells[i].getCellFormat() == defaultDateFormat))
          {
            // The cell has no overriding format specified, so
            // set it to the column default
            ((WritableCell) cells[i]).setCellFormat(xfr);
          }
        }
      }
    }

    // Write out the auto filter
    if (autoFilter != null)
    {
      autoFilter.write(outputFile);
    }

    DimensionRecord dr = new DimensionRecord(numRows, numCols);
    outputFile.write(dr);

    // Write out all the rows, in blocks of 32
    for (int block = 0; block < numBlocks; block++)
    {
      DBCellRecord dbcell = new DBCellRecord(outputFile.getPos());

      int blockRows = Math.min(32, numRows - block * 32);
      boolean firstRow = true;

      // First write out all the row records
      for (int i = block * 32; i < block * 32 + blockRows; i++)
      {
        if (rows[i] != null)
        {
          rows[i].write(outputFile);
          if (firstRow)
          {
            dbcell.setCellOffset(outputFile.getPos());
            firstRow = false;
          }
        }
      }

      // Now write out all the cells
      for (int i = block * 32; i < block * 32 + blockRows; i++)
      {
        if (rows[i] != null)
        {
          dbcell.addCellRowPosition(outputFile.getPos());
          rows[i].writeCells(outputFile);
        }
      }

      // Now set the current file position in the index record
      indexRecord.addBlockPosition(outputFile.getPos());
      
      // Set the position of the file pointer and write out the DBCell
      // record
      dbcell.setPosition(outputFile.getPos());
      outputFile.write(dbcell);
    }
    
    // Do the drawings and charts if enabled
    if (!workbookSettings.getDrawingsDisabled())
    {
      drawingWriter.write(outputFile);
    }

    Window2Record w2r = new Window2Record(settings);
    outputFile.write(w2r);

    // Handle the frozen panes
    if (settings.getHorizontalFreeze() != 0 ||
        settings.getVerticalFreeze() != 0)
    {
      PaneRecord pr = new PaneRecord(settings.getHorizontalFreeze(),
                                     settings.getVerticalFreeze());
      outputFile.write(pr);

      // Handle the selection record.  First, there will always be a top left
      SelectionRecord sr = new SelectionRecord
        (SelectionRecord.upperLeft, 0, 0);
      outputFile.write(sr);

      // Top right
      if (settings.getHorizontalFreeze() != 0)
      {
        sr = new SelectionRecord
          (SelectionRecord.upperRight, settings.getHorizontalFreeze(), 0);
        outputFile.write(sr);
      }

      // Bottom left
      if (settings.getVerticalFreeze() != 0)
      {
        sr = new SelectionRecord
          (SelectionRecord.lowerLeft, 0, settings.getVerticalFreeze());
        outputFile.write(sr);
      }

      // Bottom right
      if (settings.getHorizontalFreeze() != 0 &&
          settings.getVerticalFreeze() != 0)
      {
        sr = new SelectionRecord
          (SelectionRecord.lowerRight, 
           settings.getHorizontalFreeze(), 
           settings.getVerticalFreeze());
        outputFile.write(sr);
      }

      Weird1Record w1r = new Weird1Record();
      outputFile.write(w1r);
    }
    else
    {
      // No frozen panes - just write out the selection record for the 
      // whole sheet
      SelectionRecord sr = new SelectionRecord
        (SelectionRecord.upperLeft, 0, 0);
      outputFile.write(sr);
    }

    // Handle the zoom factor
    if (settings.getZoomFactor() != 100)
    {
      SCLRecord sclr = new SCLRecord(settings.getZoomFactor());
      outputFile.write(sclr);
    }

    // Now write out all the merged cells
    mergedCells.write(outputFile);

    // Write out all the hyperlinks
    Iterator hi = hyperlinks.iterator();
    WritableHyperlink hlr = null;
    while (hi.hasNext())
    {
      hlr = (WritableHyperlink) hi.next();
      outputFile.write(hlr);
    }

    if (buttonPropertySet != null)
    {
      outputFile.write(buttonPropertySet);
    }

    // Write out the data validations
    if (dataValidation != null || validatedCells.size() > 0)
    {
      writeDataValidation();
    }

    // Write out the conditional formats
    if (conditionalFormats != null && conditionalFormats.size() > 0)
    {
      for (Iterator i = conditionalFormats.iterator() ; i.hasNext() ; )
      {
        ConditionalFormat cf = (ConditionalFormat) i.next();
        cf.write(outputFile);
      }
    }

    EOFRecord eof = new EOFRecord();
    outputFile.write(eof);

    // Now the various cross reference offsets have been calculated,
    // retrospectively set the values in the output file
    outputFile.setData(indexRecord.getData(), indexPos+4);
  }

  /**
   * Gets the header.  Called when copying sheets
   *
   * @return the page header
   */
  final HeaderRecord getHeader()
  {
    return header;
  }

  /**
   * Gets the footer.  Called when copying sheets
   *
   * @return the page footer
   */
  final FooterRecord getFooter()
  {
    return footer;
  }

  /**
   * Sets the data necessary for writing out the sheet.  This method must
   *  be called immediately prior to writing
   *
   * @param rws the rows in the spreadsheet
   */
  void setWriteData(RowRecord[] rws, 
                    ArrayList   rb,
                    ArrayList   cb,
                    ArrayList   hl,
                    MergedCells mc,
                    TreeSet     cf,
                    int         mrol,
                    int         mcol)
  {
    rows = rws;
    rowBreaks = rb;
    columnBreaks = cb;
    hyperlinks = hl;
    mergedCells = mc;
    columnFormats = cf;
    maxRowOutlineLevel = mrol;
    maxColumnOutlineLevel = mcol;
  }

  /**
   * Sets the dimensions of this spreadsheet.  This method must be called 
   * immediately prior to writing
   *
   * @param rws the number of rows
   * @param cls the number of columns
   */
  void setDimensions(int rws, int cls)
  {
    numRows = rws;
    numCols = cls;
  }

  /**
   * Sets the sheet settings for this particular sheet.  Must be
   * called immediately prior to writing
   * 
   * @param sr the sheet settings
   */
  void setSettings(SheetSettings sr)
  {
    settings = sr;
  }

  /**
   * Accessor for the workspace options
   *
   * @return the workspace options
   */
  WorkspaceInformationRecord getWorkspaceOptions()
  {
    return workspaceOptions;
  }

  /**
   * Accessor for the workspace options
   *
   * @param wo the workspace options
   */
  void setWorkspaceOptions(WorkspaceInformationRecord wo)
  {
    if (wo != null)
    {
      workspaceOptions = wo;
    }
  }


  /**
   * Sets the charts for this sheet
   *
   * @param ch the charts
   */
  void setCharts(Chart[] ch)
  {
    drawingWriter.setCharts(ch);
  }

  /**
   * Sets the drawings on this sheet
   *
   * @param dr the list of drawings
   * @param mod a modified flag
   */
  void setDrawings(ArrayList dr, boolean mod)
  {
    drawingWriter.setDrawings(dr, mod);
  }

  /**
   * Accessor for the charts on this sheet
   *
   * @return the charts
   */
  Chart[] getCharts()
  {
    return  drawingWriter.getCharts();
  }

  /**
   * Check all the merged cells for borders.  If the merge record has
   * borders, then we need to rejig the cell formats to take account of this.
   * This is called by the write method of the WritableWorkbookImpl, so that
   * any new XFRecords that are created may be written out with the others
   */
  void checkMergedBorders()
  {
    Range[] mcells = mergedCells.getMergedCells();
    ArrayList borderFormats = new ArrayList();
    for (int mci = 0 ; mci < mcells.length ; mci++)
    {
      Range range = mcells[mci];
      Cell topLeft = range.getTopLeft();
      XFRecord tlformat = (XFRecord) topLeft.getCellFormat();

      if (tlformat != null && 
          tlformat.hasBorders() == true && 
          !tlformat.isRead())
      {
        try
        {
          CellXFRecord cf1 = new CellXFRecord(tlformat);
          Cell bottomRight = range.getBottomRight();

          cf1.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
          cf1.setBorder(Border.LEFT, 
                        tlformat.getBorderLine(Border.LEFT), 
                        tlformat.getBorderColour(Border.LEFT));
          cf1.setBorder(Border.TOP,  
                        tlformat.getBorderLine(Border.TOP),
                        tlformat.getBorderColour(Border.TOP));

          if (topLeft.getRow() == bottomRight.getRow())
          {
            cf1.setBorder(Border.BOTTOM, 
                          tlformat.getBorderLine(Border.BOTTOM),
                          tlformat.getBorderColour(Border.BOTTOM));
          }

          if (topLeft.getColumn() == bottomRight.getColumn())
          {
            cf1.setBorder(Border.RIGHT, 
                          tlformat.getBorderLine(Border.RIGHT),
                          tlformat.getBorderColour(Border.RIGHT));
          }

          int index = borderFormats.indexOf(cf1);
          if (index != -1)
          {
            cf1 = (CellXFRecord) borderFormats.get(index);
          }
          else
          {
            borderFormats.add(cf1);
          }
          ( (WritableCell) topLeft).setCellFormat(cf1);

          // Handle the bottom left corner
          if (bottomRight.getRow() > topLeft.getRow())
          {
            // Handle the corner cell
            if (bottomRight.getColumn() != topLeft.getColumn())
            {
              CellXFRecord cf2 = new CellXFRecord(tlformat);
              cf2.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf2.setBorder(Border.LEFT, 
                            tlformat.getBorderLine(Border.LEFT),
                            tlformat.getBorderColour(Border.LEFT));
              cf2.setBorder(Border.BOTTOM,  
                            tlformat.getBorderLine(Border.BOTTOM),
                            tlformat.getBorderColour(Border.BOTTOM));
            
              index = borderFormats.indexOf(cf2);
              if (index != -1)
              {
                cf2 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf2);
              }

              sheet.addCell(new Blank(topLeft.getColumn(), 
                                      bottomRight.getRow(), cf2));
            }

            // Handle the cells down the left hand side (and along the
            // right too, if necessary)
            for (int i = topLeft.getRow() + 1; i < bottomRight.getRow() ;i++)
            {
              CellXFRecord cf3 = new CellXFRecord(tlformat);
              cf3.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf3.setBorder(Border.LEFT, 
                            tlformat.getBorderLine(Border.LEFT),
                            tlformat.getBorderColour(Border.LEFT));

              if (topLeft.getColumn() == bottomRight.getColumn())
              {
                cf3.setBorder(Border.RIGHT, 
                              tlformat.getBorderLine(Border.RIGHT),
                              tlformat.getBorderColour(Border.RIGHT));
              }

              index = borderFormats.indexOf(cf3);
              if (index != -1)
              {
                cf3 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf3);
              }

              sheet.addCell(new Blank(topLeft.getColumn(), i, cf3));
            }
          }

          // Handle the top right corner
          if (bottomRight.getColumn() > topLeft.getColumn())
          {
            if (bottomRight.getRow() != topLeft.getRow())
            {
              // Handle the corner cell
              CellXFRecord cf6 = new CellXFRecord(tlformat);
              cf6.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf6.setBorder(Border.RIGHT, 
                            tlformat.getBorderLine(Border.RIGHT),
                            tlformat.getBorderColour(Border.RIGHT));
              cf6.setBorder(Border.TOP,  
                            tlformat.getBorderLine(Border.TOP),
                            tlformat.getBorderColour(Border.TOP));
              index = borderFormats.indexOf(cf6);
              if (index != -1)
              {
                cf6 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf6);
              }
              
              sheet.addCell(new Blank(bottomRight.getColumn(), 
                                      topLeft.getRow(), cf6));
            }

            // Handle the cells along the right
            for (int i = topLeft.getRow() + 1; 
                     i < bottomRight.getRow() ;i++)
            {
              CellXFRecord cf7 = new CellXFRecord(tlformat);
              cf7.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf7.setBorder(Border.RIGHT, 
                            tlformat.getBorderLine(Border.RIGHT),
                            tlformat.getBorderColour(Border.RIGHT));

              index = borderFormats.indexOf(cf7);
              if (index != -1)
              {
                cf7 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf7);
              }
              
              sheet.addCell(new Blank(bottomRight.getColumn(), i, cf7));
            }

            // Handle the cells along the top, and along the bottom too
            for (int i = topLeft.getColumn() + 1; 
                     i < bottomRight.getColumn() ;i++)
            {
              CellXFRecord cf8 = new CellXFRecord(tlformat);
              cf8.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf8.setBorder(Border.TOP, 
                            tlformat.getBorderLine(Border.TOP),
                            tlformat.getBorderColour(Border.TOP));
              
              if (topLeft.getRow() == bottomRight.getRow())
              {
                cf8.setBorder(Border.BOTTOM, 
                              tlformat.getBorderLine(Border.BOTTOM),
                              tlformat.getBorderColour(Border.BOTTOM));
              }

              index = borderFormats.indexOf(cf8);
              if (index != -1)
              {
                cf8 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf8);
              }
              
              sheet.addCell(new Blank(i, topLeft.getRow(), cf8));
            }
          }

          // Handle the bottom right corner
          if (bottomRight.getColumn() > topLeft.getColumn() ||
              bottomRight.getRow() > topLeft.getRow())
          {
            // Handle the corner cell
            CellXFRecord cf4 = new CellXFRecord(tlformat);
            cf4.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
            cf4.setBorder(Border.RIGHT,  
                          tlformat.getBorderLine(Border.RIGHT),
                          tlformat.getBorderColour(Border.RIGHT));
            cf4.setBorder(Border.BOTTOM, 
                          tlformat.getBorderLine(Border.BOTTOM),
                          tlformat.getBorderColour(Border.BOTTOM));

            if (bottomRight.getRow() == topLeft.getRow())
            {
              cf4.setBorder(Border.TOP, 
                            tlformat.getBorderLine(Border.TOP),
                            tlformat.getBorderColour(Border.TOP));
            }

            if (bottomRight.getColumn() == topLeft.getColumn())
            {
              cf4.setBorder(Border.LEFT, 
                            tlformat.getBorderLine(Border.LEFT),
                            tlformat.getBorderColour(Border.LEFT));
            }

            index = borderFormats.indexOf(cf4);
            if (index != -1)
            {
              cf4 = (CellXFRecord) borderFormats.get(index);
            }
            else
            {
              borderFormats.add(cf4);
            }

            sheet.addCell(new Blank(bottomRight.getColumn(), 
                                    bottomRight.getRow(), cf4));

            // Handle the cells along the bottom (and along the top
            // as well, if appropriate)
            for (int i = topLeft.getColumn() + 1; 
                     i < bottomRight.getColumn() ;i++)
            {
              CellXFRecord cf5 = new CellXFRecord(tlformat);
              cf5.setBorder(Border.ALL, BorderLineStyle.NONE, Colour.BLACK);
              cf5.setBorder(Border.BOTTOM, 
                            tlformat.getBorderLine(Border.BOTTOM),
                            tlformat.getBorderColour(Border.BOTTOM));

              if (topLeft.getRow() == bottomRight.getRow())
              {
                cf5.setBorder(Border.TOP, 
                              tlformat.getBorderLine(Border.TOP),
                              tlformat.getBorderColour(Border.TOP));
              }

              index = borderFormats.indexOf(cf5);
              if (index != -1)
              {
                cf5 = (CellXFRecord) borderFormats.get(index);
              }
              else
              {
                borderFormats.add(cf5);
              }
              
              sheet.addCell(new Blank(i, bottomRight.getRow(), cf5));
            }
          }
        }
        catch (WriteException e)
        {
          // just log e.toString(), not the whole stack trace
          logger.warn(e.toString());  
        }
      }
    }
  }

  /**
   * Get the cells in the column.  Don't use the interface method
   * getColumn for this as this will create loads of empty cells,
   * and we could do without that overhead
   */
  private Cell[] getColumn(int col)
  {
    // Find the last non-null cell
    boolean found = false;
    int row = numRows - 1;

    while (row >= 0 && !found)
    {
      if (rows[row] != null &&
          rows[row].getCell(col) != null)
      {
        found = true;
      }
      else
      {
        row--;
      }
    }

    // Only create entries for non-empty cells
    Cell[] cells = new Cell[row+1];

    for (int i = 0; i <= row; i++)
    {
      cells[i] = rows[i] != null ? rows[i].getCell(col) : null;
    }

    return cells;
  }

  /**
   * Sets a flag to indicate that this sheet contains a chart only
   */
  void setChartOnly()
  {
    chartOnly = true;
  }

  /**
   * Sets the environment specific print record
   *
   * @param pls the print record
   */
  void setPLS(PLSRecord pls)
  {
    plsRecord = pls;
  }

  /**
   * Sets the button property set record
   *
   * @param bps the button property set
   */
  void setButtonPropertySet(ButtonPropertySetRecord bps)
  {
    buttonPropertySet = bps;
  }

  /**
   * Sets the data validations
   *
   * @param dv the read-in list of data validations
   * @param vc the api manipulated set of data validations
   */
  void setDataValidation(DataValidation dv, ArrayList vc)
  {
    dataValidation = dv;
    validatedCells = vc;
  }

  /**
   * Sets the conditional formats
   *
   * @param cf the conditonal formats
   */
  void setConditionalFormats(ArrayList cf)
  {
    conditionalFormats = cf;
  }

  /**
   * Sets the auto filter
   *
   * @param af the autofilter
   */
  void setAutoFilter(AutoFilter af)
  {
    autoFilter = af;
  }

  /**
   * Writes out the data validations
   */
  private void writeDataValidation() throws IOException
  {
    if (dataValidation != null && validatedCells.size() == 0)
    {
      // the only data validations are those read in - this should
      // never be the case now that shared data validations add
      // to the validatedCells list
      dataValidation.write(outputFile); 
      return;
    }

    if (dataValidation == null && validatedCells.size() > 0)
    {
      // the only data validations are those which have been added by the
      // write API.  Need to sort out the combo box id
      int comboBoxId = sheet.getComboBox() != null ? 
        sheet.getComboBox().getObjectId() : DataValidation.DEFAULT_OBJECT_ID;
      dataValidation = new DataValidation(comboBoxId,
                                          sheet.getWorkbook(),
                                          sheet.getWorkbook(),
                                          workbookSettings);
    }

    for (Iterator i = validatedCells.iterator(); i.hasNext(); )
    {
      CellValue cv = (CellValue) i.next();
      CellFeatures cf = cv.getCellFeatures();

      // Do not do anything if the DVParser has been copied, as it
      // will already by on the DataValidation record as a result
      // of the SheetCopier process
      if (!cf.getDVParser().copied())
      {
        if (!cf.getDVParser().extendedCellsValidation())
        {
          // DVParser is specific for a single cell validation - just add it
          DataValiditySettingsRecord dvsr = 
            new DataValiditySettingsRecord(cf.getDVParser());
          dataValidation.add(dvsr);
        }
        else
        {
          // Only add the DVParser once for shared validations
          // only add it if it is the top left cell 
          if (cv.getColumn() == cf.getDVParser().getFirstColumn() &&
              cv.getRow()    == cf.getDVParser().getFirstRow())
          {
            DataValiditySettingsRecord dvsr = 
              new DataValiditySettingsRecord(cf.getDVParser());
            dataValidation.add(dvsr);
          }
        }
      }
    }
    dataValidation.write(outputFile);
  }
  /*

    // There is a mixture of read and write validations
    for (Iterator i = validatedCells.iterator(); i.hasNext(); )
    {
    CellValue cv = (CellValue) i.next();
    CellFeatures cf = cv.getCellFeatures();
    DataValiditySettingsRecord dvsr = 
    new DataValiditySettingsRecord(cf.getDVParser());
    dataValidation.add(dvsr);
    }
    dataValidation.write(outputFile);
    return;
    }
    */
}
