/*********************************************************************
*
*      Copyright (C) 2001 Andrew Khan
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
 License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl.write.biff;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.CellFeatures;
import jxl.CellReferenceHelper;
import jxl.Sheet;
import jxl.biff.DataValidation;
import jxl.biff.DataValiditySettingsRecord;
import jxl.biff.DVParser;
import jxl.biff.FormattingRecords;
import jxl.biff.IntegerHelper;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.biff.XFRecord;
import jxl.biff.drawing.ComboBox;
import jxl.biff.drawing.Comment;
import jxl.format.CellFormat;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableWorkbook;

/**
 * Abstract class which stores the jxl.common.data used for cells, such
 * as row, column and formatting information.  
 * Any record which directly represents the contents of a cell, such
 * as labels and numbers, are derived from this class
 * data store
 */
public abstract class CellValue extends WritableRecordData 
  implements WritableCell
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CellValue.class);
  
  /**
   * The row in the worksheet at which this cell is located
   */
  private int row;

  /**
   * The column in the worksheet at which this cell is located
   */
  private int column;

  /**
   * The format applied to this cell
   */
  private XFRecord format;
  
  /**
   * A handle to the formatting records, used in case we want
   * to change the format of the cell once it has been added
   * to the spreadsheet
   */
  private FormattingRecords formattingRecords;

  /**
   * A flag to indicate that this record is already referenced within
   * a worksheet
   */
  private boolean referenced;

  /**
   * A handle to the sheet
   */
  private WritableSheetImpl sheet;

  /**
   * The cell features
   */
  private WritableCellFeatures features;

  /**
   * Internal copied flag, to prevent cell features being added multiple
   * times to the drawing array
   */
  private boolean copied;

  /**
   * Constructor used when building writable cells from the Java API
   * 
   * @param c the column
   * @param t the type indicator
   * @param r the row
   */
  protected CellValue(Type t, int c, int r)
  {
    this(t, c, r, WritableWorkbook.NORMAL_STYLE);
    copied = false;
  }

  /**
   * Constructor used when creating a writable cell from a read-only cell 
   * (when copying a workbook)
   * 
   * @param c the cell to clone
   * @param t the type of this cell
   */
  protected CellValue(Type t, Cell c)
  {
    this(t, c.getColumn(), c.getRow());
    copied = true;

    format = (XFRecord) c.getCellFormat();
    
    if (c.getCellFeatures() != null)
    {
      features = new WritableCellFeatures(c.getCellFeatures());
      features.setWritableCell(this);
    }
  }

  /**
   * Overloaded constructor used when building writable cells from the 
   * Java API which also takes a format
   * 
   * @param c the column
   * @param t the cell type
   * @param r the row
   * @param st the format to apply to this cell
   */
  protected CellValue(Type t, int c, int r, CellFormat st)
  {
    super(t);
    row    = r;
    column = c;
    format = (XFRecord) st;
    referenced = false;
    copied = false;
  }

  /**
   * Copy constructor 
   * 
   * @param c the column
   * @param t the cell type
   * @param r the row
   * @param cv the value to copy
   */
  protected CellValue(Type t, int c, int r, CellValue cv)
  {
    super(t);
    row    = r;
    column = c;
    format = cv.format;
    referenced = false;
    copied = false; // used during a deep copy, so the cell features need 
                    // to be added again

    if (cv.features != null)
    {
      features = new WritableCellFeatures(cv.features);
      features.setWritableCell(this);
    }
  }

  /**
   * An API function which sets the format to apply to this cell
   * 
   * @param cf the format to apply to this cell
   */
  public void setCellFormat(CellFormat cf)
  {
    format = (XFRecord) cf;

    // If the referenced flag has not been set, this cell has not
    // been added to the spreadsheet, so we don't need to perform
    // any further logic
    if (!referenced)
    {
      return;
    }

    // The cell has already been added to the spreadsheet, so the 
    // formattingRecords reference must be initialized
    Assert.verify(formattingRecords != null);

    addCellFormat();
  }

  /**
   * Returns the row number of this cell
   * 
   * @return the row number of this cell
   */
  public int getRow()
  {
    return row;
  }

  /**
   * Returns the column number of this cell
   * 
   * @return the column number of this cell
   */
  public int getColumn()
  {
    return column;
  }

  /**
   * Indicates whether or not this cell is hidden, by virtue of either
   * the entire row or column being collapsed
   *
   * @return TRUE if this cell is hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    ColumnInfoRecord cir = sheet.getColumnInfo(column);
    
    if (cir != null && cir.getWidth() == 0)
    {
      return true;
    }

    RowRecord rr = sheet.getRowInfo(row);

    if (rr != null && (rr.getRowHeight() == 0 || rr.isCollapsed()))
    {
      return true;
    }

    return false;
  }

  /**
   * Gets the data to write to the output file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] mydata = new byte[6];
    IntegerHelper.getTwoBytes(row, mydata, 0);
    IntegerHelper.getTwoBytes(column, mydata, 2);
    IntegerHelper.getTwoBytes(format.getXFIndex(), mydata, 4);
    return mydata;
  }

  /**
   * Called when the cell is added to the worksheet in order to indicate
   * that this object is already added to the worksheet
   * This method also verifies that the associated formats and formats
   * have been initialized correctly
   * 
   * @param fr the formatting records
   * @param ss the shared strings used within the workbook
   * @param s the sheet this is being added to
   */
  void setCellDetails(FormattingRecords fr, SharedStrings ss, 
                      WritableSheetImpl s)
  {
    referenced = true;
    sheet = s;
    formattingRecords = fr;

    addCellFormat();
    addCellFeatures();
  }

  /**
   * Internal method to see if this cell is referenced within the workbook.
   * Once this has been placed in the workbook, it becomes immutable
   * 
   * @return TRUE if this cell has been added to a sheet, FALSE otherwise
   */
  final boolean isReferenced()
  {
    return referenced;
  }

  /**
   * Gets the internal index of the formatting record
   * 
   * @return the index of the format record
   */
  final int getXFIndex()
  {
    return format.getXFIndex();
  }

  /**
   * API method which gets the format applied to this cell
   * 
   * @return the format for this cell
   */
  public CellFormat getCellFormat()
  {
    return format;
  }

  /**
   * Increments the row of this cell by one.  Invoked by the sheet when 
   * inserting rows
   */
  void incrementRow()
  {
    row++;

    if (features != null)
    {
      Comment c = features.getCommentDrawing();
      if (c != null)
      {
        c.setX(column);
        c.setY(row);
      }
    }
  }

  /**
   * Decrements the row of this cell by one.  Invoked by the sheet when 
   * removing rows
   */
  void decrementRow()
  {
    row--;

    if (features != null)
    {
      Comment c = features.getCommentDrawing();
      if ( c!= null)
      {
        c.setX(column);
        c.setY(row);
      }

      if (features.hasDropDown())
      {
        logger.warn("need to change value for drop down drawing");
      }
    }
  }

  /**
   * Increments the column of this cell by one.  Invoked by the sheet when 
   * inserting columns
   */
  void incrementColumn()
  {
    column++;

    if (features != null)
    {
      Comment c = features.getCommentDrawing();
      if (c != null)
      {
        c.setX(column);
        c.setY(row);
      }
    }

  }

  /**
   * Decrements the column of this cell by one.  Invoked by the sheet when 
   * removing columns
   */
  void decrementColumn()
  {
    column--;

    if (features != null)
    {
      Comment c = features.getCommentDrawing();
      if (c != null)
      {
        c.setX(column);
        c.setY(row);
      }
    }

  }

  /**
   * Called when a column is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param col the column number which was inserted
   */
  void columnInserted(Sheet s, int sheetIndex, int col)
  {
  }

  /**
   * Called when a column is removed on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param col the column number which was inserted
   */
  void columnRemoved(Sheet s, int sheetIndex, int col)
  {
  }

  /**
   * Called when a row is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param row the column number which was inserted
   */
  void rowInserted(Sheet s, int sheetIndex, int row)
  {
  }

  /**
   * Called when a row is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the row was removed
   * @param sheetIndex the sheet index on which the column was removed
   * @param row the column number which was removed
   */
  void rowRemoved(Sheet s, int sheetIndex, int row)
  {
  }

  /**
   * Accessor for the sheet containing this cell
   *
   * @return the sheet containing this cell
   */
  public WritableSheetImpl getSheet()
  {
    return sheet;
  }

  /**
   * Adds the format information to the shared records.  Performs the necessary
   * checks (and clones) to ensure that the formats are not shared.
   * Called from setCellDetails and setCellFormat
   */
  private void addCellFormat()
  {
    // Check to see if the format is one of the shared Workbook defaults.  If
    // so, then get hold of the Workbook's specific instance
    Styles styles = sheet.getWorkbook().getStyles();
    format = styles.getFormat(format);

    try
    {      
      if (!format.isInitialized())
      {
        formattingRecords.addStyle(format);
      }
    }
    catch (NumFormatRecordsException e)
    {
      logger.warn("Maximum number of format records exceeded.  Using " +
                  "default format.");
      format = styles.getNormalStyle();
    }
  }

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public CellFeatures getCellFeatures()
  {
    return features;
  }

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public WritableCellFeatures getWritableCellFeatures()
  {
    return features;
  }

  /**
   * Sets the cell features
   *
   * @param cf the cell features
   */
  public void setCellFeatures(WritableCellFeatures cf)
  {
    if (features != null) 
    {
      logger.warn("current cell features for " + 
                  CellReferenceHelper.getCellReference(this) + 
                  " not null - overwriting");

      // Check to see if the features include a shared data validation
      if (features.hasDataValidation() &&
          features.getDVParser() != null &&
          features.getDVParser().extendedCellsValidation())
      {
        DVParser dvp = features.getDVParser();
        logger.warn("Cannot add cell features to " + 
                    CellReferenceHelper.getCellReference(this) + 
                    " because it is part of the shared cell validation " +
                    "group " +
                    CellReferenceHelper.getCellReference(dvp.getFirstColumn(),
                                                         dvp.getFirstRow()) +
                    "-" +
                    CellReferenceHelper.getCellReference(dvp.getLastColumn(),
                                                         dvp.getLastRow()));
        return;
      }
    }

    features = cf;
    cf.setWritableCell(this);

    // If the cell is already on the worksheet, then add the cell features
    // to the workbook
    if (referenced)
    {
      addCellFeatures();
    }
  }

  /**
   * Handles any addition cell features, such as comments or data 
   * validation.  Called internally from this class when a cell is 
   * added to the workbook, and also externally from BaseCellFeatures
   * following a call to setComment
   */
  public final void addCellFeatures()
  {
    if (features == null)
    {
      return;
    }

    if (copied == true)
    {
      copied = false;

      return;
    }

    if (features.getComment() != null)
    {
      Comment comment = new Comment(features.getComment(), 
                                    column, row);
      comment.setWidth(features.getCommentWidth());
      comment.setHeight(features.getCommentHeight());
      sheet.addDrawing(comment);      
      sheet.getWorkbook().addDrawing(comment);
      features.setCommentDrawing(comment);
    }

    if (features.hasDataValidation())
    {
      try
      {
        features.getDVParser().setCell(column, 
                                       row, 
                                       sheet.getWorkbook(), 
                                       sheet.getWorkbook(),
                                       sheet.getWorkbookSettings());
      }
      catch (jxl.biff.formula.FormulaException e)
      {
        Assert.verify(false);
      }

      sheet.addValidationCell(this);
      if (!features.hasDropDown())
      {
        return;
      }
      
      // Get the combo box drawing object for list validations
      if (sheet.getComboBox() == null)
      {
        // Need to add the combo box the first time, since even though
        // it doesn't need a separate Sp entry, it still needs to increment
        // the shape id
        ComboBox cb = new ComboBox();
        sheet.addDrawing(cb);
        sheet.getWorkbook().addDrawing(cb);
        sheet.setComboBox(cb);
      }

      features.setComboBox(sheet.getComboBox());
    }
  }

  /**
   * Internal function invoked by WritableSheetImpl called when shared data 
   * validation is removed
   */
  public final void removeCellFeatures()
  {
    /*
    // Remove the comment
    features.removeComment();

    // Remove the data validation
    features.removeDataValidation();
    */

    features = null;
  }


  /**
   * Called by the cell features to remove a comment
   *
   * @param c the comment to remove
   */
  public final void removeComment(Comment c)
  {
    sheet.removeDrawing(c);
  }

  /**
   * Called by the cell features to remove the data validation
   */
  public final void removeDataValidation()
  {
    sheet.removeDataValidation(this);
  }

  /**
   * Called when doing a copy of a writable object to indicate the source
   * was writable than a read only copy and certain things (most notably
   * the comments will need to be re-evaluated)
   *
   * @param boolean the copied flag
   */
  final void setCopied(boolean c)
  {
    copied = c;
  }

}
