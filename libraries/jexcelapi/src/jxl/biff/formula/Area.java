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

package jxl.biff.formula;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.biff.CellReferenceHelper;
import jxl.biff.IntegerHelper;

/**
 * A nested class to hold range information
 */
class Area extends Operand implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Area.class);

  /**
   * The first column
   */
  private int columnFirst;

  /**
   * The first row
   */
  private int rowFirst;

  /**
   * The last column
   */
  private int columnLast;

  /**
   * The last row
   */
  private int rowLast;

  /**
   * Indicates whether the first column is a relative or absolute reference
   */
  private boolean columnFirstRelative;

  /**
   * Indicates whether the first row is a relative or absolute reference
   */
  private boolean rowFirstRelative;

  /**
   * Indicates whether the last column is a relative or absolute reference
   */
  private boolean columnLastRelative;

  /**
   * Indicates whether the last row is a relative or absolute reference
   */
  private boolean rowLastRelative;

  /**
   * Constructor
   */
  Area()
  {
  }

  /**
   * Constructor invoked when parsing a string formula
   *
   * @param s the string to parse
   */
  Area(String s)
  {
    int seppos = s.indexOf(":");
    Assert.verify(seppos != -1);
    String startcell = s.substring(0, seppos);
    String endcell = s.substring(seppos + 1);

    columnFirst = CellReferenceHelper.getColumn(startcell);
    rowFirst = CellReferenceHelper.getRow(startcell);
    columnLast = CellReferenceHelper.getColumn(endcell);
    rowLast = CellReferenceHelper.getRow(endcell);

    columnFirstRelative = CellReferenceHelper.isColumnRelative(startcell);
    rowFirstRelative = CellReferenceHelper.isRowRelative(startcell);
    columnLastRelative = CellReferenceHelper.isColumnRelative(endcell);
    rowLastRelative = CellReferenceHelper.isRowRelative(endcell);
  }

  /**
   * Accessor for the first column
   *
   * @return the first column
   */
  int getFirstColumn()
  {
    return columnFirst;
  }

  /**
   * Accessor for the first row
   *
   * @return the first row
   */
  int getFirstRow()
  {
    return rowFirst;
  }

  /**
   * Accessor for the last column
   *
   * @return the last column
   */
  int getLastColumn()
  {
    return columnLast;
  }

  /**
   * Accessor for the last row
   *
   * @return the last row
   */
  int getLastRow()
  {
    return rowLast;
  }

  /**
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos)
  {
    rowFirst =  IntegerHelper.getInt(data[pos], data[pos + 1]);
    rowLast  = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
    int columnMask = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
    columnFirst = columnMask & 0x00ff;
    columnFirstRelative = ((columnMask & 0x4000) != 0);
    rowFirstRelative = ((columnMask & 0x8000) != 0);
    columnMask = IntegerHelper.getInt(data[pos + 6], data[pos + 7]);
    columnLast = columnMask & 0x00ff;
    columnLastRelative = ((columnMask & 0x4000) != 0);
    rowLastRelative = ((columnMask & 0x8000) != 0);

    return 8;
  }

  /**
   * Gets the string representation of this item
   *
   * @param buf the string buffer
   */
  public void getString(StringBuffer buf)
  {
    CellReferenceHelper.getCellReference(columnFirst, rowFirst, buf);
    buf.append(':');
    CellReferenceHelper.getCellReference(columnLast, rowLast, buf);
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[9];
    data[0] = !useAlternateCode() ? Token.AREA.getCode() :
                                    Token.AREA.getCode2();

    IntegerHelper.getTwoBytes(rowFirst, data, 1);
    IntegerHelper.getTwoBytes(rowLast, data, 3);

    int grcol = columnFirst;

    // Set the row/column relative bits if applicable
    if (rowFirstRelative)
    {
      grcol |= 0x8000;
    }

    if (columnFirstRelative)
    {
      grcol |= 0x4000;
    }

    IntegerHelper.getTwoBytes(grcol, data, 5);

    grcol = columnLast;

    // Set the row/column relative bits if applicable
    if (rowLastRelative)
    {
      grcol |= 0x8000;
    }

    if (columnLastRelative)
    {
      grcol |= 0x4000;
    }

    IntegerHelper.getTwoBytes(grcol, data, 7);

    return data;
  }

  /**
   * Adjusts all the relative cell references in this formula by the
   * amount specified.  Used when copying formulas
   *
   * @param colAdjust the amount to add on to each relative cell reference
   * @param rowAdjust the amount to add on to each relative row reference
   */
  public void adjustRelativeCellReferences(int colAdjust, int rowAdjust)
  {
    if (columnFirstRelative)
    {
      columnFirst += colAdjust;
    }

    if (columnLastRelative)
    {
      columnLast += colAdjust;
    }

    if (rowFirstRelative)
    {
      rowFirst += rowAdjust;
    }

    if (rowLastRelative)
    {
      rowLast += rowAdjust;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param col the column number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void columnInserted(int sheetIndex, int col, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (col <= columnFirst)
    {
      columnFirst++;
    }

    if (col <= columnLast)
    {
      columnLast++;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was removed
   * @param col the column number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void columnRemoved(int sheetIndex, int col, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (col < columnFirst)
    {
      columnFirst--;
    }

    if (col <= columnLast)
    {
      columnLast--;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was inserted
   * @param row the row number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowInserted(int sheetIndex, int row, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (rowLast == 0xffff)
    {
      // area applies to the whole column, so nothing to do
      return;
    }

    if (row <= rowFirst)
    {
      rowFirst++;
    }

    if (row <= rowLast)
    {
      rowLast++;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was removed
   * @param row the row number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowRemoved(int sheetIndex, int row, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (rowLast == 0xffff)
    {
      // area applies to the whole column, so nothing to do
      return;
    }

    if (row < rowFirst)
    {
      rowFirst--;
    }

    if (row <= rowLast)
    {
      rowLast--;
    }
  }

  /**
   * Used by subclasses columns/row range to set the range information
   *
   * @param colFirst the first column
   * @param colLast the last column
   * @param rwFirst the first row
   * @param rwLast the last row
   * @param colFirstRel flag indicating whether the first column is relative
   * @param colLastRel flag indicating whether the last column is relative
   * @param rowFirstRel flag indicating whether the first row is relative
   * @param rowLastRel flag indicating whether the last row is relative
   */
  protected void setRangeData(int colFirst,
                              int colLast,
                              int rwFirst,
                              int rwLast,
                              boolean colFirstRel,
                              boolean colLastRel,
                              boolean rowFirstRel,
                              boolean rowLastRel)
  {
    columnFirst = colFirst;
    columnLast = colLast;
    rowFirst = rwFirst;
    rowLast = rwLast;
    columnFirstRelative = colFirstRel;
    columnLastRelative = colLastRel;
    rowFirstRelative = rowFirstRel;
    rowLastRelative = rowLastRel;
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing here
   */
  void handleImportedCellReferences()
  {
  }
}
