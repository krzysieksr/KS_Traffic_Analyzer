/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

package jxl.biff;

import jxl.common.Logger;

import jxl.read.biff.Record;

/**
 * Range information for conditional formatting
 */
public class ConditionalFormatRangeRecord extends WritableRecordData
{
  // The logger
  private static Logger logger = 
    Logger.getLogger(ConditionalFormatRangeRecord.class);

  /**
   * The enclosing range
   */
  private Range enclosingRange;

  /**
   * The discrete ranges
   */
  private Range[] ranges;

  /**
   * The number of ranges
   */
  private int numRanges;

  /**
   * Initialized flag
   */
  private boolean initialized;

  /**
   * Modified flag
   */
  private boolean modified;

  /**
   * The data
   */
  private byte[] data;

  private static class Range
  {
    public int firstRow;
    public int firstColumn;
    public int lastRow;
    public int lastColumn;
    public boolean modified;

    public Range()
    {
      modified = false;
    }

    /**
     * Inserts a blank column into this spreadsheet.  If the column is out of 
     * range of the columns in the sheet, then no action is taken
     *
     * @param col the column to insert
     */
    public void insertColumn(int col)
    {
      if (col > lastColumn)
      {
        return;
      }

      if (col <= firstColumn)
      {
        firstColumn++;
        modified = true;
      }

      if (col <= lastColumn)
      {
        lastColumn++;
        modified = true;
      }
    }

    /**
     * Removes a column from this spreadsheet.  If the column is out of range
     * of the columns in the sheet, then no action is taken
     *
     * @param col the column to remove
     */
    public void removeColumn(int col)
    {
      if (col > lastColumn)
      {
        return;
      }

      if (col < firstColumn)
      {
        firstColumn--;
        modified = true;
      }

      if (col <= lastColumn)
      {
        lastColumn--;
        modified = true;
      }
    }

    /**
     * Removes a row from this spreadsheet.  If the row is out of 
     * range of the columns in the sheet, then no action is taken
     *
     * @param row the row to remove
     */
    public void removeRow(int row)
    {
      if (row > lastRow)
      {
        return;
      }

      if (row < firstRow)
      {
        firstRow--;
        modified = true;
      }

      if (row <= lastRow)
      {
        lastRow--;
        modified = true;
      }
    }

    /**
     * Inserts a blank row into this spreadsheet.  If the row is out of range
     * of the rows in the sheet, then no action is taken
     *
     * @param row the row to insert
     */
    public void insertRow(int row)
    {
      if (row > lastRow)
      {
        return;
      }

      if (row <= firstRow)
      {
        firstRow++;
        modified = true;
      }

      if (row <= lastRow)
      {
        lastRow++;
        modified = true;
      }
    }

  }

  /**
   * Constructor
   */
  public ConditionalFormatRangeRecord(Record t)
  {
    super(t);

    initialized = false;
    modified = false;
    data = getRecord().getData();
  }

  /**
   * Initialization function
   */
  private void initialize()
  {
    enclosingRange = new Range();
    enclosingRange.firstRow = IntegerHelper.getInt(data[4], data[5]);
    enclosingRange.lastRow = IntegerHelper.getInt(data[6], data[7]);
    enclosingRange.firstColumn = IntegerHelper.getInt(data[8], data[9]);    
    enclosingRange.lastColumn = IntegerHelper.getInt(data[10], data[11]);
    numRanges = IntegerHelper.getInt(data[12], data[13]);
    ranges = new Range[numRanges];

    int pos = 14;

    for (int i = 0; i < numRanges; i++)
    {
      ranges[i] = new Range();
      ranges[i].firstRow = IntegerHelper.getInt(data[pos], data[pos+1]);
      ranges[i].lastRow = IntegerHelper.getInt(data[pos+2], data[pos+3]);
      ranges[i].firstColumn = IntegerHelper.getInt(data[pos+4], data[pos+5]);
      ranges[i].lastColumn = IntegerHelper.getInt(data[pos+6], data[pos+7]);
      pos += 8;
    }

    initialized = true;
  }

  /**
   * Inserts a blank column into this spreadsheet.  If the column is out of 
   * range of the columns in the sheet, then no action is taken
   *
   * @param col the column to insert
   */
  public void insertColumn(int col)
  {
    if (!initialized)
    {
      initialize();
    }

    enclosingRange.insertColumn(col);
    if (enclosingRange.modified)
    {
      modified = true;
    }

    for (int i = 0 ; i < ranges.length ; i++)
    {
      ranges[i].insertColumn(col);

      if (ranges[i].modified)
      {
        modified = true;
      }
    }

    return;
  }

  /**
   * Inserts a blank column into this spreadsheet.  If the column is out of 
   * range of the columns in the sheet, then no action is taken
   *
   * @param col the column to insert
   */
  public void removeColumn(int col)
  {
    if (!initialized)
    {
      initialize();
    }

    enclosingRange.removeColumn(col);
    if (enclosingRange.modified)
    {
      modified = true;
    }

    for (int i = 0 ; i < ranges.length ; i++)
    {
      ranges[i].removeColumn(col);

      if (ranges[i].modified)
      {
        modified = true;
      }
    }

    return;
  }

  /**
   * Removes a row from this spreadsheet.  If the row is out of 
   * range of the columns in the sheet, then no action is taken
   *
   * @param row the row to remove
   */
  public void removeRow(int row)
  {
    if (!initialized)
    {
      initialize();
    }

    enclosingRange.removeRow(row);
    if (enclosingRange.modified)
    {
      modified = true;
    }

    for (int i = 0 ; i < ranges.length ; i++)
    {
      ranges[i].removeRow(row);

      if (ranges[i].modified)
      {
        modified = true;
      }
    }

    return;
  }

  /**
   * Inserts a blank row into this spreadsheet.  If the row is out of range
   * of the rows in the sheet, then no action is taken
   *
   * @param row the row to insert
   */
  public void insertRow(int row)
  {
    if (!initialized)
    {
      initialize();
    }

    enclosingRange.insertRow(row);
    if (enclosingRange.modified)
    {
      modified = true;
    }

    for (int i = 0 ; i < ranges.length ; i++)
    {
      ranges[i].insertRow(row);

      if (ranges[i].modified)
      {
        modified = true;
      }
    }

    return;
  }


  /**
   * Retrieves the data for output to binary file
   * 
   * @return the data to be written
   */
  public byte[] getData()
  {
    if (!modified)
    {
      return data;
    }

    byte[] d = new byte[14 + ranges.length * 8];

    // Copy in the original information
    System.arraycopy(data, 0, d, 0, 4);

    // Create the new range
    IntegerHelper.getTwoBytes(enclosingRange.firstRow, d, 4);
    IntegerHelper.getTwoBytes(enclosingRange.lastRow, d, 6);
    IntegerHelper.getTwoBytes(enclosingRange.firstColumn, d, 8);
    IntegerHelper.getTwoBytes(enclosingRange.lastColumn, d, 10);

    IntegerHelper.getTwoBytes(numRanges, d, 12);

    int pos = 14;
    for (int i = 0 ; i < ranges.length ; i++)
    {
      IntegerHelper.getTwoBytes(ranges[i].firstRow, d, pos);
      IntegerHelper.getTwoBytes(ranges[i].lastRow, d, pos+2);
      IntegerHelper.getTwoBytes(ranges[i].firstColumn, d, pos+4);
      IntegerHelper.getTwoBytes(ranges[i].lastColumn, d, pos+6);
      pos += 8;
    }

    return d;
  }
}

