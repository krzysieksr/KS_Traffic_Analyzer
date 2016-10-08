/*********************************************************************
*
*      Copyright (C) 2006 Andrew Khan
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

import java.util.Arrays;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.TreeSet;
import java.util.Iterator;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.CellView;
import jxl.DateCell;
import jxl.HeaderFooter;
import jxl.Hyperlink;
import jxl.Image;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Range;
import jxl.Sheet;
import jxl.SheetSettings;
import jxl.WorkbookSettings;
import jxl.biff.AutoFilter;
import jxl.biff.CellReferenceHelper;
import jxl.biff.ConditionalFormat;
import jxl.biff.DataValidation;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.IndexMapping;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.SheetRangeImpl;
import jxl.biff.XFRecord;
import jxl.biff.drawing.Chart;
import jxl.biff.drawing.CheckBox;
import jxl.biff.drawing.ComboBox;
import jxl.biff.drawing.Drawing;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.format.CellFormat;
import jxl.biff.formula.FormulaException;
import jxl.read.biff.SheetImpl;
import jxl.read.biff.NameRecord;
import jxl.read.biff.WorkbookParser;
import jxl.write.Blank;
import jxl.write.Boolean;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableHyperlink;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * A transient utility object used to copy sheets.   This 
 * functionality has been farmed out to a different class
 * in order to reduce the bloat of the WritableSheetImpl
 */
class SheetCopier
{
  private static Logger logger = Logger.getLogger(SheetCopier.class);

  private SheetImpl fromSheet;
  private WritableSheetImpl toSheet;
  private WorkbookSettings workbookSettings;

  // Objects used by the sheet
  private TreeSet columnFormats;
  private FormattingRecords formatRecords;
  private ArrayList hyperlinks;
  private MergedCells mergedCells;
  private ArrayList rowBreaks;
  private ArrayList columnBreaks;
  private SheetWriter sheetWriter;
  private ArrayList drawings;
  private ArrayList images;
  private ArrayList conditionalFormats;
  private ArrayList validatedCells;
  private AutoFilter autoFilter;
  private DataValidation dataValidation;
  private ComboBox comboBox;
  private PLSRecord plsRecord;
  private boolean chartOnly;
  private ButtonPropertySetRecord buttonPropertySet;
  private int numRows;
  private int maxRowOutlineLevel;
  private int maxColumnOutlineLevel;

  // Objects used to maintain state during the copy process
  private HashMap xfRecords;
  private HashMap fonts;
  private HashMap formats;

  public SheetCopier(Sheet f, WritableSheet t)
  {
    fromSheet = (SheetImpl) f;
    toSheet = (WritableSheetImpl) t;
    workbookSettings = toSheet.getWorkbook().getSettings();
    chartOnly = false;
  }

  void setColumnFormats(TreeSet cf)
  {
    columnFormats = cf;
  }

  void setFormatRecords(FormattingRecords fr)
  {
    formatRecords = fr;
  }

  void setHyperlinks(ArrayList h)
  {
    hyperlinks = h;
  }

  void setMergedCells(MergedCells mc)
  {
    mergedCells = mc;
  }

  void setRowBreaks(ArrayList rb)
  {
    rowBreaks = rb;
  }

  void setColumnBreaks(ArrayList cb)
  {
    columnBreaks = cb;
  }

  void setSheetWriter(SheetWriter sw)
  {
    sheetWriter = sw;
  }

  void setDrawings(ArrayList d)
  {
    drawings = d;
  }

  void setImages(ArrayList i)
  {
    images = i;
  }

  void setConditionalFormats(ArrayList cf)
  {
    conditionalFormats = cf;
  }

  void setValidatedCells(ArrayList vc)
  {
    validatedCells = vc;
  }

  AutoFilter getAutoFilter()
  {
    return autoFilter;
  }

  DataValidation getDataValidation()
  {
    return dataValidation;
  }

  ComboBox getComboBox()
  {
    return comboBox;
  }

  PLSRecord getPLSRecord()
  {
    return plsRecord;
  }

  boolean isChartOnly()
  {
    return chartOnly;
  }

  ButtonPropertySetRecord getButtonPropertySet()
  {
    return buttonPropertySet;
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copySheet()
  {
    shallowCopyCells();

    // Copy the column info records
    jxl.read.biff.ColumnInfoRecord[] readCirs = fromSheet.getColumnInfos();

    for (int i = 0 ; i < readCirs.length; i++)
    {
      jxl.read.biff.ColumnInfoRecord rcir = readCirs[i];
      for (int j = rcir.getStartColumn(); j <= rcir.getEndColumn() ; j++) 
      {
        ColumnInfoRecord cir = new ColumnInfoRecord(rcir, j, 
                                                    formatRecords);
        cir.setHidden(rcir.getHidden());
        columnFormats.add(cir);
      }
    }

    // Copy the hyperlinks
    Hyperlink[] hls = fromSheet.getHyperlinks();
    for (int i = 0 ; i < hls.length; i++)
    {
      WritableHyperlink hr = new WritableHyperlink
        (hls[i], toSheet);
      hyperlinks.add(hr);
    }

    // Copy the merged cells
    Range[] merged = fromSheet.getMergedCells();

    for (int i = 0; i < merged.length; i++)
    {
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl)merged[i], toSheet));
    }

    // Copy the row properties
    try
    {
      jxl.read.biff.RowRecord[] rowprops  = fromSheet.getRowProperties();

      for (int i = 0; i < rowprops.length; i++)
      {
        RowRecord rr = toSheet.getRowRecord(rowprops[i].getRowNumber());
        XFRecord format = rowprops[i].hasDefaultFormat() ? 
          formatRecords.getXFRecord(rowprops[i].getXFIndex()) : null;
        rr.setRowDetails(rowprops[i].getRowHeight(), 
                         rowprops[i].matchesDefaultFontHeight(),
                         rowprops[i].isCollapsed(),
                         rowprops[i].getOutlineLevel(),
                         rowprops[i].getGroupStart(),
                         format);
        numRows = Math.max(numRows, rowprops[i].getRowNumber() + 1);
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the headers and footers
    //    sheetWriter.setHeader(new HeaderRecord(si.getHeader()));
    //    sheetWriter.setFooter(new FooterRecord(si.getFooter()));

    // Copy the page breaks
    int[] rowbreaks = fromSheet.getRowPageBreaks();

    if (rowbreaks != null)
    {
      for (int i = 0; i < rowbreaks.length; i++)
      {
        rowBreaks.add(new Integer(rowbreaks[i]));
      }
    }

    int[] columnbreaks = fromSheet.getColumnPageBreaks();

    if (columnbreaks != null)
    {
      for (int i = 0; i < columnbreaks.length; i++)
      {
        columnBreaks.add(new Integer(columnbreaks[i]));
      }
    }

    // Copy the charts
    sheetWriter.setCharts(fromSheet.getCharts());

    // Copy the drawings
    DrawingGroupObject[] dr = fromSheet.getDrawings();
    for (int i = 0 ; i < dr.length ; i++)
    {
      if (dr[i] instanceof jxl.biff.drawing.Drawing)
      {
        WritableImage wi = new WritableImage
          (dr[i], toSheet.getWorkbook().getDrawingGroup());
        drawings.add(wi);
        images.add(wi);
      }
      else if (dr[i] instanceof jxl.biff.drawing.Comment)
      {
        jxl.biff.drawing.Comment c = 
          new jxl.biff.drawing.Comment(dr[i], 
                                       toSheet.getWorkbook().getDrawingGroup(),
                                       workbookSettings);
        drawings.add(c);
        
        // Set up the reference on the cell value
        CellValue cv = (CellValue) toSheet.getWritableCell(c.getColumn(), 
                                                           c.getRow());
        Assert.verify(cv.getCellFeatures() != null);
        cv.getWritableCellFeatures().setCommentDrawing(c);
      }
      else if (dr[i] instanceof jxl.biff.drawing.Button)
      {
        jxl.biff.drawing.Button b = 
          new jxl.biff.drawing.Button
          (dr[i], 
           toSheet.getWorkbook().getDrawingGroup(),
           workbookSettings);
        drawings.add(b);
      }
      else if (dr[i] instanceof jxl.biff.drawing.ComboBox)
      {
        jxl.biff.drawing.ComboBox cb = 
          new jxl.biff.drawing.ComboBox
          (dr[i], 
           toSheet.getWorkbook().getDrawingGroup(), 
           workbookSettings);
        drawings.add(cb);
      }
      else if (dr[i] instanceof jxl.biff.drawing.CheckBox)
      {
        jxl.biff.drawing.CheckBox cb = 
          new jxl.biff.drawing.CheckBox
          (dr[i], 
           toSheet.getWorkbook().getDrawingGroup(), 
           workbookSettings);
        drawings.add(cb);
      }

    }

    // Copy the data validations
    DataValidation rdv = fromSheet.getDataValidation();
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv, 
                                          toSheet.getWorkbook(), 
                                          toSheet.getWorkbook(),
                                          workbookSettings);
      int objid = dataValidation.getComboBoxObjectId();

      if (objid != 0)
      {
        comboBox = (ComboBox) drawings.get(objid);
      }
    }

    // Copy the conditional formats
    ConditionalFormat[] cf = fromSheet.getConditionalFormats();
    if (cf.length > 0)
    {
      for (int i = 0; i < cf.length ; i++)
      {
        conditionalFormats.add(cf[i]);
      }
    }

    // Get the autofilter
    autoFilter = fromSheet.getAutoFilter();

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromSheet.getWorkspaceOptions());

    // Set a flag to indicate if it contains a chart only
    if (fromSheet.getSheetBof().isChart())
    {
      chartOnly = true;
      sheetWriter.setChartOnly();
    }

    // Copy the environment specific print record
    if (fromSheet.getPLS() != null)
    {
      if (fromSheet.getWorkbookBof().isBiff7())
      {
        logger.warn("Cannot copy Biff7 print settings record - ignoring");
      }
      else
      {
        plsRecord = new PLSRecord(fromSheet.getPLS());
      }
    }

    // Copy the button property set
    if (fromSheet.getButtonPropertySet() != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromSheet.getButtonPropertySet());
    }

    // Copy the outline levels
    maxRowOutlineLevel = fromSheet.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = fromSheet.getMaxColumnOutlineLevel();
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copyWritableSheet()
  {
    shallowCopyCells();

    /*
    // Copy the column formats
    Iterator cfit = fromWritableSheet.columnFormats.iterator();
    while (cfit.hasNext())
    {
      ColumnInfoRecord cv = new ColumnInfoRecord
        ((ColumnInfoRecord) cfit.next());
      columnFormats.add(cv);
    }

    // Copy the merged cells
    Range[] merged = fromWritableSheet.getMergedCells();

    for (int i = 0; i < merged.length; i++)
    {
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl)merged[i], this));
    }

    // Copy the row properties
    try
    {
      RowRecord[] copyRows = fromWritableSheet.rows;
      RowRecord row = null;
      for (int i = 0; i < copyRows.length ; i++)
      {
        row = copyRows[i];
        
        if (row != null &&
            (!row.isDefaultHeight() ||
             row.isCollapsed()))
        {
          RowRecord rr = getRowRecord(i);
          rr.setRowDetails(row.getRowHeight(), 
                           row.matchesDefaultFontHeight(),
                           row.isCollapsed(), 
                           row.getStyle());
        }
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the horizontal page breaks
    rowBreaks = new ArrayList(fromWritableSheet.rowBreaks);

    // Copy the vertical page breaks
    columnBreaks = new ArrayList(fromWritableSheet.columnBreaks);

    // Copy the data validations
    DataValidation rdv = fromWritableSheet.dataValidation;
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv, 
                                          workbook,
                                          workbook, 
                                          workbookSettings);
    }

    // Copy the charts 
    sheetWriter.setCharts(fromWritableSheet.getCharts());

    // Copy the drawings
    DrawingGroupObject[] dr = si.getDrawings();
    for (int i = 0 ; i < dr.length ; i++)
    {
      if (dr[i] instanceof jxl.biff.drawing.Drawing)
      {
        WritableImage wi = new WritableImage(dr[i], 
                                             workbook.getDrawingGroup());
        drawings.add(wi);
        images.add(wi);
      }

      // Not necessary to copy the comments, as they will be handled by
      // the deep copy of the individual cells
    }

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromWritableSheet.getWorkspaceOptions());

    // Copy the environment specific print record
    if (fromWritableSheet.plsRecord != null)
    {
      plsRecord = new PLSRecord(fromWritableSheet.plsRecord);
    }

    // Copy the button property set
    if (fromWritableSheet.buttonPropertySet != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromWritableSheet.buttonPropertySet);
    }
    */
  }  

  /**
   * Imports a sheet from a different workbook, doing a deep copy
   */
  public void importSheet()
  {
    xfRecords = new HashMap();
    fonts = new HashMap();
    formats = new HashMap();

    deepCopyCells();

    // Copy the column info records
    jxl.read.biff.ColumnInfoRecord[] readCirs = fromSheet.getColumnInfos();

    for (int i = 0 ; i < readCirs.length; i++)
    {
      jxl.read.biff.ColumnInfoRecord rcir = readCirs[i];
      for (int j = rcir.getStartColumn(); j <= rcir.getEndColumn() ; j++) 
      {
        ColumnInfoRecord cir = new ColumnInfoRecord(rcir, j);
        int xfIndex = cir.getXfIndex();
        XFRecord cf = (WritableCellFormat) xfRecords.get(new Integer(xfIndex));

        if (cf == null)
        {
          CellFormat readFormat = fromSheet.getColumnView(j).getFormat();
          WritableCellFormat wcf = copyCellFormat(readFormat);
        }

        cir.setCellFormat(cf);
        cir.setHidden(rcir.getHidden());
        columnFormats.add(cir);
      }
    }

    // Copy the hyperlinks
    Hyperlink[] hls = fromSheet.getHyperlinks();
    for (int i = 0 ; i < hls.length; i++)
    {
      WritableHyperlink hr = new WritableHyperlink
        (hls[i], toSheet);
      hyperlinks.add(hr);
    }

    // Copy the merged cells
    Range[] merged = fromSheet.getMergedCells();

    for (int i = 0; i < merged.length; i++)
    {
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl)merged[i], toSheet));
    }

    // Copy the row properties
    try
    {
      jxl.read.biff.RowRecord[] rowprops  = fromSheet.getRowProperties();

      for (int i = 0; i < rowprops.length; i++)
      {
        RowRecord rr = toSheet.getRowRecord(rowprops[i].getRowNumber());
        XFRecord format = null;
        jxl.read.biff.RowRecord rowrec = rowprops[i];
        if (rowrec.hasDefaultFormat())
        {
          format = (WritableCellFormat) 
            xfRecords.get(new Integer(rowrec.getXFIndex()));

          if (format == null)
          {
            int rownum = rowrec.getRowNumber();
            CellFormat readFormat = fromSheet.getRowView(rownum).getFormat();
            WritableCellFormat wcf = copyCellFormat(readFormat);
          }
        }

        rr.setRowDetails(rowrec.getRowHeight(), 
                         rowrec.matchesDefaultFontHeight(),
                         rowrec.isCollapsed(),
                         rowrec.getOutlineLevel(),
                         rowrec.getGroupStart(),
                         format);
        numRows = Math.max(numRows, rowprops[i].getRowNumber() + 1);
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the headers and footers
    //    sheetWriter.setHeader(new HeaderRecord(si.getHeader()));
    //    sheetWriter.setFooter(new FooterRecord(si.getFooter()));

    // Copy the page breaks
    int[] rowbreaks = fromSheet.getRowPageBreaks();

    if (rowbreaks != null)
    {
      for (int i = 0; i < rowbreaks.length; i++)
      {
        rowBreaks.add(new Integer(rowbreaks[i]));
      }
    }

    int[] columnbreaks = fromSheet.getColumnPageBreaks();

    if (columnbreaks != null)
    {
      for (int i = 0; i < columnbreaks.length; i++)
      {
        columnBreaks.add(new Integer(columnbreaks[i]));
      }
    }

    // Copy the charts
    Chart[] fromCharts = fromSheet.getCharts();
    if (fromCharts != null && fromCharts.length > 0)
    {
      logger.warn("Importing of charts is not supported");
      /*
      sheetWriter.setCharts(fromSheet.getCharts());
      IndexMapping xfMapping = new IndexMapping(200);
      for (Iterator i = xfRecords.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        XFRecord xfmapping = (XFRecord) xfRecords.get(key);
        xfMapping.setMapping(key.intValue(), xfmapping.getXFIndex());
      }

      IndexMapping fontMapping = new IndexMapping(200);
      for (Iterator i = fonts.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        Integer fontmap = (Integer) fonts.get(key);
        fontMapping.setMapping(key.intValue(), fontmap.intValue());
      }

      IndexMapping formatMapping = new IndexMapping(200);
      for (Iterator i = formats.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        Integer formatmap = (Integer) formats.get(key);
        formatMapping.setMapping(key.intValue(), formatmap.intValue());
      }

      // Now reuse the rationalization feature on each chart  to
      // handle the new fonts
      for (int i = 0; i < fromCharts.length ; i++)
      {
        fromCharts[i].rationalize(xfMapping, fontMapping, formatMapping);
      }
      */
    }

    // Copy the drawings
    DrawingGroupObject[] dr = fromSheet.getDrawings();

    // Make sure the destination workbook has a drawing group
    // created in it
    if (dr.length > 0 && 
        toSheet.getWorkbook().getDrawingGroup() == null)
    {
      toSheet.getWorkbook().createDrawingGroup();
    }

    for (int i = 0 ; i < dr.length ; i++)
    {
      if (dr[i] instanceof jxl.biff.drawing.Drawing)
      {
        WritableImage wi = new WritableImage
          (dr[i].getX(), dr[i].getY(), 
           dr[i].getWidth(), dr[i].getHeight(),
           dr[i].getImageData());
        toSheet.getWorkbook().addDrawing(wi);
        drawings.add(wi);
        images.add(wi);
      }
      else if (dr[i] instanceof jxl.biff.drawing.Comment)
      {
        jxl.biff.drawing.Comment c = 
          new jxl.biff.drawing.Comment(dr[i], 
                                       toSheet.getWorkbook().getDrawingGroup(),
                                       workbookSettings);
        drawings.add(c);
        
        // Set up the reference on the cell value
        CellValue cv = (CellValue) toSheet.getWritableCell(c.getColumn(), 
                                                           c.getRow());
        Assert.verify(cv.getCellFeatures() != null);
        cv.getWritableCellFeatures().setCommentDrawing(c);
      }
      else if (dr[i] instanceof jxl.biff.drawing.Button)
      {
        jxl.biff.drawing.Button b = 
          new jxl.biff.drawing.Button
          (dr[i], 
           toSheet.getWorkbook().getDrawingGroup(),
           workbookSettings);
        drawings.add(b);
      }
      else if (dr[i] instanceof jxl.biff.drawing.ComboBox)
      {
        jxl.biff.drawing.ComboBox cb = 
          new jxl.biff.drawing.ComboBox
          (dr[i], 
           toSheet.getWorkbook().getDrawingGroup(), 
           workbookSettings);
        drawings.add(cb);
      }
    }

    // Copy the data validations
    DataValidation rdv = fromSheet.getDataValidation();
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv, 
                                          toSheet.getWorkbook(), 
                                          toSheet.getWorkbook(),
                                          workbookSettings);
      int objid = dataValidation.getComboBoxObjectId();
      if (objid != 0)
      {
        comboBox = (ComboBox) drawings.get(objid);
      }
    }

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromSheet.getWorkspaceOptions());

    // Set a flag to indicate if it contains a chart only
    if (fromSheet.getSheetBof().isChart())
    {
      chartOnly = true;
      sheetWriter.setChartOnly();
    }

    // Copy the environment specific print record
    if (fromSheet.getPLS() != null)
    {
      if (fromSheet.getWorkbookBof().isBiff7())
      {
        logger.warn("Cannot copy Biff7 print settings record - ignoring");
      }
      else
      {
        plsRecord = new PLSRecord(fromSheet.getPLS());
      }
    }

    // Copy the button property set
    if (fromSheet.getButtonPropertySet() != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromSheet.getButtonPropertySet());
    }

    importNames();

    // Copy the outline levels
    maxRowOutlineLevel = fromSheet.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = fromSheet.getMaxColumnOutlineLevel();
  }

  /**
   * Performs a shallow copy of the specified cell
   */
  private WritableCell shallowCopyCell(Cell cell)
  {
    CellType ct = cell.getType();
    WritableCell newCell = null;

    if (ct == CellType.LABEL)
    {
      newCell = new Label((LabelCell) cell);
    }
    else if (ct == CellType.NUMBER)
    {
      newCell = new Number((NumberCell) cell);
    }
    else if (ct == CellType.DATE)
    {
      newCell = new DateTime((DateCell) cell);
    }
    else if (ct == CellType.BOOLEAN)
    {
      newCell = new Boolean((BooleanCell) cell);
    }
    else if (ct == CellType.NUMBER_FORMULA)
    {
      newCell = new ReadNumberFormulaRecord((FormulaData) cell);
    }
    else if (ct == CellType.STRING_FORMULA)
    {
      newCell = new ReadStringFormulaRecord((FormulaData) cell);
    }
    else if( ct == CellType.BOOLEAN_FORMULA)
    {
      newCell = new ReadBooleanFormulaRecord((FormulaData) cell);
    }
    else if (ct == CellType.DATE_FORMULA)
    {
      newCell = new ReadDateFormulaRecord((FormulaData) cell);
    }
    else if(ct == CellType.FORMULA_ERROR)
    {
      newCell = new ReadErrorFormulaRecord((FormulaData) cell);
    }
    else if (ct == CellType.EMPTY)
    {
      if (cell.getCellFormat() != null)
      {
        // It is a blank cell, rather than an empty cell, so
        // it may have formatting information, so
        // it must be copied
        newCell = new Blank(cell);
      }
    }
    
    return newCell;
  }

  /** 
   * Performs a deep copy of the specified cell, handling the cell format
   * 
   * @param cell the cell to copy
   */
  private WritableCell deepCopyCell(Cell cell)
  {
    WritableCell c = shallowCopyCell(cell);

    if (c == null)
    {
      return c;
    }

    if (c instanceof ReadFormulaRecord)
    {
      ReadFormulaRecord rfr = (ReadFormulaRecord) c;
      boolean crossSheetReference = !rfr.handleImportedCellReferences
        (fromSheet.getWorkbook(),
         fromSheet.getWorkbook(),
         workbookSettings);
      
      if (crossSheetReference)
      {
        try
        {
        logger.warn("Formula " + rfr.getFormula() +
                    " in cell " + 
                    CellReferenceHelper.getCellReference(cell.getColumn(),
                                                         cell.getRow()) +
                    " cannot be imported because it references another " +
                    " sheet from the source workbook");
        }
        catch (FormulaException e)
        {
          logger.warn("Formula  in cell " + 
                      CellReferenceHelper.getCellReference(cell.getColumn(),
                                                           cell.getRow()) +
                      " cannot be imported:  " + e.getMessage());
        }
        
        // Create a new error formula and add it instead
        c = new Formula(cell.getColumn(), cell.getRow(), "\"ERROR\"");
      }
    }

    // Copy the cell format
    CellFormat cf = c.getCellFormat();
    int index = ( (XFRecord) cf).getXFIndex();
    WritableCellFormat wcf = (WritableCellFormat) 
      xfRecords.get(new Integer(index));

    if (wcf == null)
    {
      wcf = copyCellFormat(cf);
    }

    c.setCellFormat(wcf);

    return c;
  }

  /** 
   * Perform a shallow copy of the cells from the specified sheet into this one
   */
  void shallowCopyCells()
  {
    // Copy the cells
    int cells = fromSheet.getRows();
    Cell[] row = null;
    Cell cell = null;
    for (int i = 0;  i < cells; i++)
    {
      row = fromSheet.getRow(i);

      for (int j = 0; j < row.length; j++)
      {
        cell = row[j];
        WritableCell c = shallowCopyCell(cell);

        // Encase the calls to addCell in a try-catch block
        // These should not generate any errors, because we are
        // copying from an existing spreadsheet.  In the event of
        // errors, catch the exception and then bomb out with an
        // assertion
        try
        {
          if (c != null)
          {
            toSheet.addCell(c);

            // Cell.setCellFeatures short circuits when the cell is copied,
            // so make sure the copy logic handles the validated cells        
            if (c.getCellFeatures() != null &&
                c.getCellFeatures().hasDataValidation())
            {
              validatedCells.add(c);
            }
          }
        }
        catch (WriteException e)
        {
          Assert.verify(false);
        }
      }
    }
    numRows = toSheet.getRows();
  }

  /** 
   * Perform a deep copy of the cells from the specified sheet into this one
   */
  void deepCopyCells()
  {
    // Copy the cells
    int cells = fromSheet.getRows();
    Cell[] row = null;
    Cell cell = null;
    for (int i = 0;  i < cells; i++)
    {
      row = fromSheet.getRow(i);

      for (int j = 0; j < row.length; j++)
      {
        cell = row[j];
        WritableCell c = deepCopyCell(cell);

        // Encase the calls to addCell in a try-catch block
        // These should not generate any errors, because we are
        // copying from an existing spreadsheet.  In the event of
        // errors, catch the exception and then bomb out with an
        // assertion
        try
        {
          if (c != null)
          {
            toSheet.addCell(c);

            // Cell.setCellFeatures short circuits when the cell is copied,
            // so make sure the copy logic handles the validated cells
            if (c.getCellFeatures() != null &
                c.getCellFeatures().hasDataValidation())
            {
              validatedCells.add(c);
            }
          }
        }
        catch (WriteException e)
        {
          Assert.verify(false);
        }
      }
    }
  }

  /**
   * Returns an initialized copy of the cell format
   *
   * @param cf the cell format to copy
   * @return a deep copy of the cell format
   */
  private WritableCellFormat copyCellFormat(CellFormat cf)
  {
    try
    {
      // just do a deep copy of the cell format for now.  This will create
      // a copy of the format and font also - in the future this may
      // need to be sorted out
      XFRecord xfr = (XFRecord) cf;
      WritableCellFormat f = new WritableCellFormat(xfr);
      formatRecords.addStyle(f);

      // Maintain the local list of formats
      int xfIndex = xfr.getXFIndex();
      xfRecords.put(new Integer(xfIndex), f);

      int fontIndex = xfr.getFontIndex();
      fonts.put(new Integer(fontIndex), new Integer(f.getFontIndex()));

      int formatIndex = xfr.getFormatRecord();
      formats.put(new Integer(formatIndex), new Integer(f.getFormatRecord()));

      return f;
    }
    catch (NumFormatRecordsException e)
    {
      logger.warn("Maximum number of format records exceeded.  Using " +
                  "default format.");

      return WritableWorkbook.NORMAL_STYLE;
    }
  }

  /**
   * Imports any names defined on the source sheet to the destination workbook
   */
  private void importNames()
  {
    WorkbookParser fromWorkbook = (WorkbookParser) fromSheet.getWorkbook();
    WritableWorkbook toWorkbook = toSheet.getWorkbook();
    int fromSheetIndex = fromWorkbook.getIndex(fromSheet);
    NameRecord[] nameRecords = fromWorkbook.getNameRecords();
    String[] names = toWorkbook.getRangeNames();

    for (int i = 0 ; i < nameRecords.length ;i++)
    {
      NameRecord.NameRange[] nameRanges = nameRecords[i].getRanges();
      
      for (int j = 0; j < nameRanges.length; j++)
      {
        int nameSheetIndex = fromWorkbook.getExternalSheetIndex
          (nameRanges[j].getExternalSheet());

        if (fromSheetIndex == nameSheetIndex)
        {
          String name = nameRecords[i].getName();
          if (Arrays.binarySearch(names, name) < 0)
          {
            toWorkbook.addNameArea(name,
                                   toSheet,
                                   nameRanges[j].getFirstColumn(),
                                   nameRanges[j].getFirstRow(),
                                   nameRanges[j].getLastColumn(),
                                   nameRanges[j].getLastRow());
          }
          else
          {
            logger.warn("Named range " + name + 
                        " is already present in the destination workbook");
          }
                                 
        }
      }
    }
  }

  /**
   * Gets the number of rows - allows for the case where formatting has
   * been applied to rows, even though the row has no data
   *
   * @return the number of rows
   */
  int getRows()
  {
    return numRows;
  }

  /** 
   * Accessor for the maximum column outline level
   *
   * @return the maximum column outline level, or 0 if no outlines/groups
   */
  public int getMaxColumnOutlineLevel() 
  {
    return maxColumnOutlineLevel;
  }

  /** 
   * Accessor for the maximum row outline level
   *
   * @return the maximum row outline level, or 0 if no outlines/groups
   */
  public int getMaxRowOutlineLevel() 
  {
    return maxRowOutlineLevel;
  }
}
