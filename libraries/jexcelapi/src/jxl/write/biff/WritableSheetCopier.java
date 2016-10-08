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
import jxl.biff.WorkspaceInformationRecord;
import jxl.biff.XFRecord;
import jxl.biff.drawing.Chart;
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
class WritableSheetCopier
{
  private static Logger logger = Logger.getLogger(SheetCopier.class);

  private WritableSheetImpl fromSheet;
  private WritableSheetImpl toSheet;
  private WorkbookSettings workbookSettings;

  // Objects used by the sheet
  private TreeSet fromColumnFormats;
  private TreeSet toColumnFormats;
  private MergedCells fromMergedCells;
  private MergedCells toMergedCells;
  private RowRecord[] fromRows;
  private ArrayList fromRowBreaks;
  private ArrayList fromColumnBreaks;
  private ArrayList toRowBreaks;
  private ArrayList toColumnBreaks;
  private DataValidation fromDataValidation;
  private DataValidation toDataValidation;
  private SheetWriter sheetWriter;
  private ArrayList fromDrawings;
  private ArrayList toDrawings;
  private ArrayList toImages;
  private WorkspaceInformationRecord fromWorkspaceOptions;
  private PLSRecord fromPLSRecord;
  private PLSRecord toPLSRecord;
  private ButtonPropertySetRecord fromButtonPropertySet;
  private ButtonPropertySetRecord toButtonPropertySet;
  private ArrayList fromHyperlinks;
  private ArrayList toHyperlinks;
  private ArrayList validatedCells;
  private int numRows;
  private int maxRowOutlineLevel;
  private int maxColumnOutlineLevel;


  private boolean chartOnly;
  private FormattingRecords formatRecords;



  // Objects used to maintain state during the copy process
  private HashMap xfRecords;
  private HashMap fonts;
  private HashMap formats;

  public WritableSheetCopier(WritableSheet f, WritableSheet t)
  {
    fromSheet = (WritableSheetImpl) f;
    toSheet = (WritableSheetImpl) t;
    workbookSettings = toSheet.getWorkbook().getSettings();
    chartOnly = false;
  }

  void setColumnFormats(TreeSet fcf, TreeSet tcf)
  {
    fromColumnFormats = fcf;
    toColumnFormats = tcf;
  }

  void setMergedCells(MergedCells fmc, MergedCells tmc)
  {
    fromMergedCells = fmc;
    toMergedCells = tmc;
  }

  void setRows(RowRecord[] r)
  {
    fromRows = r;
  }

  void setValidatedCells(ArrayList vc)
  {
    validatedCells = vc;
  }

  void setRowBreaks(ArrayList frb, ArrayList trb)
  {
    fromRowBreaks = frb;
    toRowBreaks = trb;
  }

  void setColumnBreaks(ArrayList fcb, ArrayList tcb)
  {
    fromColumnBreaks = fcb;
    toColumnBreaks = tcb;
  }

  void setDrawings(ArrayList fd, ArrayList td, ArrayList ti)
  {
    fromDrawings = fd;
    toDrawings = td;
    toImages = ti;
  }

  void setHyperlinks(ArrayList fh, ArrayList th)
  {
    fromHyperlinks = fh;
    toHyperlinks = th;
  }

  void setWorkspaceOptions(WorkspaceInformationRecord wir)
  {
    fromWorkspaceOptions = wir;
  }

  void setDataValidation(DataValidation dv)
  {
    fromDataValidation = dv;
  }

  void setPLSRecord(PLSRecord plsr)
  {
    fromPLSRecord = plsr;
  }

  void setButtonPropertySetRecord(ButtonPropertySetRecord bpsr)
  {
    fromButtonPropertySet = bpsr;
  }

  void setSheetWriter(SheetWriter sw)
  {
    sheetWriter = sw;
  }


  DataValidation getDataValidation()
  {
    return toDataValidation;
  }

  PLSRecord getPLSRecord()
  {
    return toPLSRecord;
  }

  boolean isChartOnly()
  {
    return chartOnly;
  }

  ButtonPropertySetRecord getButtonPropertySet()
  {
    return toButtonPropertySet;
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copySheet()
  {
    shallowCopyCells();

    // Copy the column formats
    Iterator cfit = fromColumnFormats.iterator();
    while (cfit.hasNext())
    {
      ColumnInfoRecord cv = new ColumnInfoRecord
        ((ColumnInfoRecord) cfit.next());
      toColumnFormats.add(cv);
    }

    // Copy the merged cells
    Range[] merged = fromMergedCells.getMergedCells();

    for (int i = 0; i < merged.length; i++)
    {
      toMergedCells.add(new SheetRangeImpl((SheetRangeImpl)merged[i], 
                                           toSheet));
    }

    try
    {
      RowRecord row = null;
      RowRecord newRow = null;
      for (int i = 0; i < fromRows.length ; i++)
      {
        row = fromRows[i];
        
        if (row != null &&
            (!row.isDefaultHeight() ||
             row.isCollapsed()))
        {
          newRow = toSheet.getRowRecord(i);
          newRow.setRowDetails(row.getRowHeight(), 
                               row.matchesDefaultFontHeight(),
                               row.isCollapsed(),
                               row.getOutlineLevel(),
                               row.getGroupStart(),
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
    toRowBreaks = new ArrayList(fromRowBreaks);

    // Copy the vertical page breaks
    toColumnBreaks = new ArrayList(fromColumnBreaks);

    // Copy the data validations
    if (fromDataValidation != null)
    {
      toDataValidation = new DataValidation
        (fromDataValidation, 
         toSheet.getWorkbook(),
         toSheet.getWorkbook(),
         toSheet.getWorkbook().getSettings());
    }

    // Copy the charts
    sheetWriter.setCharts(fromSheet.getCharts());

    // Copy the drawings
    for (Iterator i = fromDrawings.iterator(); i.hasNext(); )
    {
      Object o = i.next();
      if (o instanceof jxl.biff.drawing.Drawing)
      {
        WritableImage wi = new WritableImage
          ((jxl.biff.drawing.Drawing) o, 
           toSheet.getWorkbook().getDrawingGroup());
        toDrawings.add(wi);
        toImages.add(wi);
      }

      // Not necessary to copy the comments, as they will be handled by
      // the deep copy of the individual cells
    }

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromWorkspaceOptions);

    // Copy the environment specific print record
    if (fromPLSRecord != null)
    {
      toPLSRecord = new PLSRecord(fromPLSRecord);
    }

    // Copy the button property set
    if (fromButtonPropertySet != null)
    {
      toButtonPropertySet = new ButtonPropertySetRecord(fromButtonPropertySet);
    }

    // Copy the hyperlinks
    for (Iterator i = fromHyperlinks.iterator(); i.hasNext();)
    {
      WritableHyperlink hr = new WritableHyperlink
        ((WritableHyperlink) i.next(), toSheet);
      toHyperlinks.add(hr);
    }
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
