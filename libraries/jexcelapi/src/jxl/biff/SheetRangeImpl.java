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

package jxl.biff;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;

/**
 * Implementation class for the Range interface.  This merely
 * holds the raw range information.  This implementation is used
 * for ranges which are present on the current working sheet, so the
 * getSheetIndex merely returns -1
 */
public class SheetRangeImpl implements Range
{
  /**
   * A handle to the sheet containing this range
   */
  private Sheet sheet;

  /**
   * The column number of the cell at the top left of the range
   */
  private int column1;

  /**
   * The row number of the cell at the top left of the range
   */
  private int row1;

  /**
   * The column index of the cell at the bottom right
   */
  private int column2;

  /**
   * The row index of the cell at the bottom right
   */
  private int row2;

  /**
   * Constructor
   * @param s the sheet containing the range
   * @param c1 the column number of the top left cell of the range
   * @param r1 the row number of the top left cell of the range
   * @param c2 the column number of the bottom right cell of the range
   * @param r2 the row number of the bottomr right cell of the range
   */
  public SheetRangeImpl(Sheet s, int c1, int r1,
                        int c2, int r2)
  {
    sheet   = s;
    row1    = r1;
    row2    = r2;
    column1 = c1;
    column2 = c2;
  }

  /**
   * A copy constructor used for copying ranges between sheets
   *
   * @param c the range to copy from
   * @param s the writable sheet
   */
  public SheetRangeImpl(SheetRangeImpl c, Sheet s)
  {
    sheet = s;
    row1 = c.row1;
    row2 = c.row2;
    column1 = c.column1;
    column2 = c.column2;
  }

  /**
   * Gets the cell at the top left of this range
   *
   * @return the cell at the top left
   */
  public Cell getTopLeft()
  {
    // If the print area exceeds the bounds of the sheet, then handle
    // it here.  The sheet implementation will give a NPE
    if (column1 >= sheet.getColumns() ||
        row1 >= sheet.getRows())
    {
      return new EmptyCell(column1,row1);
    }

    return sheet.getCell(column1, row1);
  }

  /**
   * Gets the cell at the bottom right of this range
   *
   * @return the cell at the bottom right
   */
  public Cell getBottomRight()
  {
    // If the print area exceeds the bounds of the sheet, then handle
    // it here.  The sheet implementation will give a NPE
    if (column2 >= sheet.getColumns() ||
        row2 >= sheet.getRows())
    {
      return new EmptyCell(column2,row2);
    }

    return sheet.getCell(column2, row2);
  }

  /**
   * Not supported.  Returns -1, indicating that it refers to the current
   * sheet
   *
   * @return -1
   */
  public int getFirstSheetIndex()
  {
    return -1;
  }

  /**
   * Not supported.  Returns -1, indicating that it refers to the current
   * sheet
   *
   * @return -1
   */
  public int getLastSheetIndex()
  {
    return -1;
  }

  /**
   * Sees whether there are any intersections between this range and the
   * range passed in.  This method is used internally by the WritableSheet to
   * verify the integrity of merged cells, hyperlinks etc.  Ranges are
   * only ever compared for the same sheet
   *
   * @param range the range to compare against
   * @return TRUE if the ranges intersect, FALSE otherwise
   */
  public boolean intersects(SheetRangeImpl range)
  {
    if (range == this)
    {
      return true;
    }

    if (row2    < range.row1    ||
        row1    > range.row2    ||
        column2 < range.column1 ||
        column1 > range.column2)
    {
      return false;
    }

    return true;
  }

  /**
   * To string method - primarily used during debugging
   *
   * @return the string version of this object
   */
  public String toString()
  {
    StringBuffer sb = new StringBuffer();
    CellReferenceHelper.getCellReference(column1, row1, sb);
    sb.append('-');
    CellReferenceHelper.getCellReference(column2, row2, sb);
    return sb.toString();
  }

  /**
   * A row has been inserted, so adjust the range objects accordingly
   *
   * @param r the row which has been inserted
   */
  public void insertRow(int r)
  {
    if (r > row2)
    {
      return;
    }

    if (r <= row1)
    {
      row1++;
    }

    if (r <= row2)
    {
      row2++;
    }
  }

  /**
   * A column has been inserted, so adjust the range objects accordingly
   *
   * @param c the column which has been inserted
   */
  public void insertColumn(int c)
  {
    if (c > column2)
    {
      return;
    }

    if (c <= column1)
    {
      column1++;
    }

    if (c <= column2)
    {
      column2++;
    }
  }

  /**
   * A row has been removed, so adjust the range objects accordingly
   *
   * @param r the row which has been inserted
   */
  public void removeRow(int r)
  {
    if (r > row2)
    {
      return;
    }

    if (r < row1)
    {
      row1--;
    }

    if (r < row2)
    {
      row2--;
    }
  }

  /**
   * A column has been removed, so adjust the range objects accordingly
   *
   * @param c the column which has been removed
   */
  public void removeColumn(int c)
  {
    if (c > column2)
    {
      return;
    }

    if (c < column1)
    {
      column1--;
    }

    if (c < column2)
    {
      column2--;
    }
  }

  /**
   * Standard hash code method
   *
   * @return the hash code
   */
  public int hashCode()
  {
    return 0xffff ^ row1 ^ row2 ^ column1 ^ column2;
  }

  /**
   * Standard equals method
   *
   * @param o the object to compare
   * @return TRUE if the two objects are the same, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof SheetRangeImpl))
    {
      return false;
    }

    SheetRangeImpl compare = (SheetRangeImpl) o;

    return (column1 == compare.column1 &&
            column2 == compare.column2 &&
            row1    == compare.row1 &&
            row2    == compare.row2);
  }

}










