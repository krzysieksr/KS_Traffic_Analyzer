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

import jxl.common.Logger;

import jxl.CellType;
import jxl.SheetSettings;
import jxl.WorkbookSettings;
import jxl.biff.CellReferenceHelper;
import jxl.biff.IndexMapping;
import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.biff.XFRecord;
import jxl.write.Number;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableSheet;

/**
 * Contains all the cells for a given row in a sheet
 */
class RowRecord extends WritableRecordData
{
  /**
   * The logger
   */
  private static final Logger logger = Logger.getLogger(RowRecord.class);

  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The cells which comprise this row
   */
  private CellValue[] cells;
  /**
   * The height of this row in 1/20ths of a point
   */
  private int rowHeight;
  /**
   * Flag to indicate whether this row is outline collapsed or not
   */
  private boolean collapsed;
  /**
   * The number of this row within the worksheet
   */
  private int rowNumber;
  /**
   * The number of columns in this row.  This is the largest column value + 1
   */
  private int numColumns;
  /**
   * The xfIndex for this row
   */
  private int xfIndex;
  /**
   * The style for this row
   */
  private XFRecord style;
  /**
   * Flag indicating that this row record has an default format
   */
  private boolean defaultFormat;
  /**
   * Flag indicating whether this row matches the default font height
   */
  private boolean matchesDefFontHeight;
  /**
   * The amount to grow the cells array by
   */
  private static final int growSize = 10;

  /**
   * The maximum integer value that can be squeezed into 30 bits
   */
  private static final int maxRKValue = 0x1fffffff;

  /**
   * The minimum integer value that can be squeezed into 30 bits
   */
  private static final int minRKValue = -0x20000000;

  /**
   * Indicates that the row is default height
   */
  private static int defaultHeightIndicator = 0xff;

  /**
   * The maximum number of columns
   */
  private static int maxColumns = 256;

  /** 
   * The outline level of the row
   */
  private int outlineLevel;

  /** 
   * Is this the icon indicator row of a group?
   */
  private boolean groupStart;
 
  /**
   * A handle back to the sheet
   */
  private WritableSheet sheet;

  /**
   * Constructs an empty row which has the specified row number
   * 
   * @param rn the row number of this row
   */
  public RowRecord(int rn, WritableSheet ws)
  {
    super(Type.ROW);
    rowNumber  = rn;
    cells      = new CellValue[0];
    numColumns  = 0;
    rowHeight  = defaultHeightIndicator;
    collapsed  = false;
    matchesDefFontHeight = true;
    sheet = ws;
  }

  /**
   * Sets the height of this row
   * 
   * @param h the row height
   */
  public void setRowHeight(int h)
  {
    if (h == 0)
    {
      setCollapsed(true);
      matchesDefFontHeight = false;
    }
    else
    {
      rowHeight = h;
      matchesDefFontHeight = false;
    }
  }

  /**
   * Sets the row details based upon the readable row record passed in
   * Called when copying spreadsheets
   *
   * @param height the height of the row record in 1/20ths of a point
   * @param mdfh matches the default font height
   * @param col the collapsed status of the row
   * @param ol the outline level
   * @param gs the group start
   * @param xf the xfrecord for the row (NULL if no default is set)
   */
  void setRowDetails(int height, 
                     boolean mdfh, 
                     boolean col, 
                     int ol, 
                     boolean gs,
                     XFRecord xfr)
  {
    rowHeight = height;
    collapsed = col;
    matchesDefFontHeight = mdfh;
    outlineLevel = ol;
    groupStart = gs;
    
    if (xfr != null)
    {
      defaultFormat = true;
      style = xfr;
      xfIndex = style.getXFIndex();
    }
  }

  /**
   * Sets the collapsed status of this row
   *
   * @param c the collapsed flag
   */
  public void setCollapsed(boolean c)
  {
    collapsed = c;
  }
  
  /**
   * Gets the row number of this row
   * 
   * @return the row number
   */
  public int getRowNumber()
  {
    return rowNumber;
  }

  /**
   * Adds a cell to this row, growing the array of cells as required
   * 
   * @param cv the cell to add
   */
  public void addCell(CellValue cv)
  {
    int col = cv.getColumn();

    if (col >= maxColumns)
    {
      logger.warn("Could not add cell at " + 
                  CellReferenceHelper.getCellReference(cv.getRow(), 
                                                       cv.getColumn()) + 
                  " because it exceeds the maximum column limit");
      return;
    }

    // Grow the array if needs be
    if (col >= cells.length)
    {
      CellValue[] oldCells = cells;
      cells = new CellValue[Math.max(oldCells.length + growSize, col+1)];
      System.arraycopy(oldCells, 0, cells, 0, oldCells.length);
      oldCells = null;
    }

    // Remove any cell features from the cell being replaced
    if (cells[col] != null)
    {
      WritableCellFeatures wcf = cells[col].getWritableCellFeatures();
      if (wcf != null)
      {
        wcf.removeComment();

        // if the cell is part of a shared data validation,then don't remove
        // the validation
        if (wcf.getDVParser() != null &&
            !wcf.getDVParser().extendedCellsValidation())
        {
          wcf.removeDataValidation();
        }
      }

    }

    cells[col] = cv;

    numColumns = Math.max(col+1, numColumns);
  }

  /**
   * Removes a cell from this row
   * 
   * @param col the column at which to remove the cell
   */
  public void removeCell(int col)
  {
    // Grow the array if needs be
    if (col >= numColumns)
    {
      return;
    }

    cells[col] = null;
  }

  /**
   * Writes out the row information data (but not the individual cells)
   * 
   * @exception IOException 
   * @param outputFile the output file
   */
  public void write(File outputFile) throws IOException
  {
    outputFile.write(this);
  }

  /**
   * Writes out all the cells in this row.  If more than three integer
   * values occur consecutively, then a MulRK record is used to group the
   * numbers
   * 
   * @exception IOException 
   * @param outputFile the output file
   */
  public void writeCells(File outputFile) 
    throws IOException
  {
    // This is the list for integer values
    ArrayList integerValues = new ArrayList();
    boolean integerValue = false;

    // Write out all the records
    for (int i = 0; i < numColumns; i++)
    {
      integerValue = false;
      if (cells[i] != null)
      {
        // See if this cell is a 30-bit integer value (without additional
        // cell features)
        if (cells[i].getType() == CellType.NUMBER)
        {
          Number nc = (Number) cells[i];
          if (nc.getValue() == (int) nc.getValue() && 
              nc.getValue() < maxRKValue &&
              nc.getValue() > minRKValue &&
              nc.getCellFeatures() == null)
          {
            integerValue = true;
          }
        }

        if (integerValue)
        {
          // This cell is an integer, add it to the list
          integerValues.add(cells[i]);
        }
        else
        {
          // This cell is not an integer.  Write out whatever integers we
          // have, and then write out this cell
          writeIntegerValues(integerValues, outputFile);
          outputFile.write(cells[i]);

          // If the cell is a string formula, write out the string record
          // immediately afterwards
          if (cells[i].getType() == CellType.STRING_FORMULA)
          {
            StringRecord sr = new StringRecord(cells[i].getContents());
            outputFile.write(sr);
          }
        }
      }
      else
      {
        // Cell does not exist.  Write out the list of integers that
        // we have
        writeIntegerValues(integerValues, outputFile);
      }
    }
    
    // All done.  Write out any remaining integer values
    writeIntegerValues(integerValues, outputFile);
  }

  /**
   * Writes out the list of integer values.  If there are more than three,
   * a MulRK record is used, otherwise a sequence of Numbers is used
   * 
   * @exception IOException 
   * @param outputFile the output file
   * @param integerValues the array of integer values
   */
  private void writeIntegerValues(ArrayList integerValues, File outputFile)
   throws IOException
  {
    if (integerValues.size() == 0)
    {
      return;
    }

    if (integerValues.size() >= 3 )
    {
      // Write out as a MulRK record
      MulRKRecord mulrk = new MulRKRecord(integerValues);
      outputFile.write(mulrk);
    }
    else
    {
      // Write out as number records
      Iterator i = integerValues.iterator();
      while (i.hasNext())
      {
        outputFile.write((CellValue) i.next());
      }
    }

    // Clear out the list of integerValues
    integerValues.clear();
  }

  /**
   * Gets the row data to output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    // Write out the row record
    byte[] data = new byte[16];

    // If the default row height has been changed in the sheet settings,
    // then we need to set the rowHeight on this row explicitly, as 
    // specifying the "match default" flag doesn't work
    int rh = rowHeight;
    if (sheet.getSettings().getDefaultRowHeight() != 
        SheetSettings.DEFAULT_DEFAULT_ROW_HEIGHT)
    {
      // the default row height has been changed.  If this row does not
      // have a specific row height set on it, then set it to the default
      if (rh == defaultHeightIndicator)
      {
        rh = sheet.getSettings().getDefaultRowHeight();
      }
    }

    IntegerHelper.getTwoBytes(rowNumber, data, 0);
    IntegerHelper.getTwoBytes(numColumns, data, 4);
    IntegerHelper.getTwoBytes(rh, data, 6);

    int options = 0x100 + outlineLevel;

    if (groupStart)
    {
      options |= 0x10;
    }

    if (collapsed)
    {
      options |= 0x20;
    }

    if (!matchesDefFontHeight)
    {
      options |= 0x40;
    }

    if (defaultFormat)
    {
      options |= 0x80;
      options |= (xfIndex << 16);
    }

    IntegerHelper.getFourBytes(options, data, 12);
    
    return data;
  }

  /**
   * Gets the maximum column value which occurs in this row
   * 
   * @return the maximum column value
   */
  public int getMaxColumn()
  {
    return numColumns;
  }

  /**
   * Gets the cell which occurs at the specified column value
   * 
   * @param col the colun for which to return the cell
   * @return the cell value at the specified position, or null if the column 
   *     is invalid
   */
  public CellValue getCell(int col)
  {
    return (col >= 0 && col < numColumns) ? cells[col] : null;
  }

  /**
   * Increments the row of this cell by one.  Invoked by the sheet when 
   * inserting rows
   */
  void incrementRow()
  {
    rowNumber++;

    for (int i = 0; i < cells.length; i++)
    {
      if (cells[i] != null)
      {
        cells[i].incrementRow();
      }
    }
  }

  /**
   * Decrements the row of this cell by one.  Invoked by the sheet when 
   * removing rows
   */
  void decrementRow()
  {
    rowNumber--;
    for (int i = 0; i < cells.length; i++)
    {
      if (cells[i] != null)
      {
        cells[i].decrementRow();
      }
    }
  }

  /**
   * Inserts a new column at the position specified.  If the max column length
   * is already reached, then the last column simply gets dropped
   *
   * @param col the column to insert
   */
  void insertColumn(int col)
  {
    // Don't bother doing anything unless there are cells after the
    // column to be inserted
    if (col >= numColumns)
    {
      return;
    }

    // Create a new array to hold the new column.  Grow it if need be
    CellValue[] oldCells = cells;

    if (numColumns  >= cells.length - 1)
    {
      cells = new CellValue[oldCells.length + growSize];
    }
    else
    {
      cells = new CellValue[oldCells.length];
    }

    // Copy in everything up to the new column
    System.arraycopy(oldCells, 0, cells, 0, col);
    
    // Copy in the remaining cells
    System.arraycopy(oldCells, col, cells, col+1, numColumns - col);

    // Increment all the internal column numbers by one
    for (int i = col+1; i <= numColumns; i++)
    {
      if (cells[i] != null)
      {
        cells[i].incrementColumn();
      }
    }

    // Adjust the maximum column record
    numColumns = Math.min(numColumns+1, maxColumns);
  }

  /**
   * Remove the new column at the position specified
   *
   * @param col the column to remove
   */
  void removeColumn(int col)
  {
    // Don't bother doing anything unless there are cells after the
    // column to be inserted
    if (col >= numColumns)
    {
      return;
    }

    // Create a new array to hold the new columns
    CellValue[] oldCells = cells;

    cells = new CellValue[oldCells.length];

    // Copy in everything up to the column
    System.arraycopy(oldCells, 0, cells, 0, col);
    
    // Copy in the remaining cells after the column
    System.arraycopy(oldCells, col + 1, cells, col, numColumns - (col+1));

    // Decrement all the internal column numbers by one
    for (int i = col; i < numColumns; i++)
    {
      if (cells[i] != null)
      {
        cells[i].decrementColumn();
      }
    }

    // Adjust the maximum column record
    numColumns--;
  }

  /**
   * Interrogates whether this row is of default height
   *
   * @return TRUE if this is set to the default height, FALSE otherwise
   */
  public boolean isDefaultHeight()
  {
    return rowHeight == defaultHeightIndicator;
  }

  /**
   * Gets the height of the row
   *
   * @return the row height
   */
  public int getRowHeight()
  {
    return rowHeight;
  }

  /**
   * Queries whether the row is collapsed
   *
   * @return the collapsed indicator
   */
  public boolean isCollapsed()
  {
    return collapsed;
  }

  /**
   * Rationalizes the sheets xf index mapping
   * @param xfmapping the index mapping
   */
  void rationalize(IndexMapping xfmapping)
  {
    if (defaultFormat)
    {
      xfIndex = xfmapping.getNewIndex(xfIndex);
    }
  }

  /**
   * Accessor for the style.  The returned value is only non-null if the
   * default style is overridden
   *
   * @return the style
   */
  XFRecord getStyle()
  {
    return style;
  }

  /**
   * Accessor for the default format flag
   *
   * @return TRUE if this row has its own default format
   */
  boolean hasDefaultFormat()
  {
    return defaultFormat;
  }

  /**
   * Accessor for the matches default font height  flag
   *
   * @return TRUE if this row matches the default font height
   */
  boolean matchesDefaultFontHeight()
  {
    return matchesDefFontHeight;
  }

  /** 
   * Accessor for the column's outline level
   *
   * @return the column's outline level
   */
  public int getOutlineLevel() 
  {
    return outlineLevel;
  }

  /** 
   * Accessor for row's groupStart state
   *
   * @return the row's groupStart state
   */
  public boolean getGroupStart() 
  {
    return groupStart;
  }

  /** 
   * Increments the row's outline level.  This is how groups are made as well
   */
  public void incrementOutlineLevel() 
  {
    outlineLevel++;
  }

  /** 
   * Decrements the row's outline level.  This removes it from a grouping 
   * level.  If
   *  all outline levels are gone the uncollapse the row.
   */
  public void decrementOutlineLevel() 
  {
    if (0 < outlineLevel)
    {
      outlineLevel--;
    }

    if (0==outlineLevel)
    {
      collapsed = false;
    }
  }

  /** 
   * Sets the row's outline level
   *
   * @param level the row's outline level
   */
  public void setOutlineLevel(int level) 
  {
    outlineLevel = level;
  }

  /**
   *  Sets the row's group start state
   *
   * @param value the group start state
   */
  public void setGroupStart(boolean value) 
  {
    groupStart = value;
  }
}










