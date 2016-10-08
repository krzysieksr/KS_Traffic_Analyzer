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

import jxl.common.Logger;

import jxl.Cell;
import jxl.biff.CellReferenceHelper;
import jxl.biff.IntegerHelper;

/**
 * A cell reference in a formula
 */
class CellReference extends Operand implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CellReference.class);

  /**
   * Indicates whether the column reference is relative or absolute
   */
  private boolean columnRelative;

  /**
   * Indicates whether the row reference is relative or absolute
   */
  private boolean rowRelative;

  /**
   * The column reference
   */
  private int column;

  /**
   * The row reference
   */
  private int row;

  /**
   * The cell containing the formula.  Stored in order to determine
   * relative cell values
   */
  private Cell relativeTo;

  /**
   * Constructor
   *
   * @param rt the cell containing the formula
   */
  public CellReference(Cell rt)
  {
    relativeTo = rt;
  }

  /**
   * Constructor
   */
  public CellReference()
  {
  }

  /**
   * Constructor invoked when parsing a text string
   *
   * @param s the string being parsed
   */
  public CellReference(String s)
  {
    column = CellReferenceHelper.getColumn(s);
    row = CellReferenceHelper.getRow(s);
    columnRelative = CellReferenceHelper.isColumnRelative(s);
    rowRelative = CellReferenceHelper.isRowRelative(s);
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
    row = IntegerHelper.getInt(data[pos], data[pos + 1]);
    int columnMask = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
    column = columnMask & 0x00ff;
    columnRelative = ((columnMask & 0x4000) != 0);
    rowRelative = ((columnMask & 0x8000) != 0);

    return 4;
  }

  /**
   * Accessor for the column
   *
   * @return the column
   */
  public int getColumn()
  {
    return column;
  }

  /**
   * Accessor for the row
   *
   * @return the row
   */
  public int getRow()
  {
    return row;
  }

  /**
   * Gets the cell reference as a string for this item
   *
   * @param buf the string buffer to populate
   */
  public void getString(StringBuffer buf)
  {
    CellReferenceHelper.getCellReference(column, !columnRelative,
                                         row, !rowRelative,
                                         buf);
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[5];
    data[0] = !useAlternateCode() ? Token.REF.getCode() :
                                    Token.REF.getCode2();

    IntegerHelper.getTwoBytes(row, data, 1);

    int grcol = column;

    // Set the row/column relative bits if applicable
    if (rowRelative)
    {
      grcol |= 0x8000;
    }

    if (columnRelative)
    {
      grcol |= 0x4000;
    }

    IntegerHelper.getTwoBytes(grcol, data, 3);

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
    if (columnRelative)
    {
      column += colAdjust;
    }

    if (rowRelative)
    {
      row += rowAdjust;
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
  public void columnInserted(int sheetIndex, int col, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (column >= col)
    {
      column++;
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

    if (column >= col)
    {
      column--;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was inserted
   * @param r the row number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowInserted(int sheetIndex, int r, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (row >= r)
    {
      row++;
    }
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was removed
   * @param r the row number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowRemoved(int sheetIndex, int r, boolean currentSheet)
  {
    if (!currentSheet)
    {
      return;
    }

    if (row >= r)
    {
      row--;
    }
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Flags the formula as invalid
   * Does nothing here
   */
  void handleImportedCellReferences()
  {
  }
}
