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
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl.write;

import jxl.NumberCell;
import jxl.format.CellFormat;
import jxl.write.biff.NumberRecord;

/**
 * A cell, created by user applications, which contains a numerical value
 */
public class Number extends NumberRecord implements WritableCell, NumberCell
{
  /**
   * Constructs a number, which, when added to a spreadsheet, will display the
   * specified value at the column/row position indicated.  By default, the
   * cell will display with an accuracy of 3 decimal places
   *
   * @param c the column
   * @param r the row
   * @param val the value
   */
  public Number(int c, int r, double val)
  {
    super(c, r, val);
  }

  /**
   * Constructs a number, which, when added to a spreadsheet, will display the
   * specified value at the column/row position with the specified CellFormat.
   * The CellFormat may specify font information and number format information
   * such as the number of decimal places
   *
   * @param c the column
   * @param r the row
   * @param val the value
   * @param st the cell format
   */
  public Number(int c, int r, double val, CellFormat st)
  {
    super(c, r, val, st);
  }

  /**
   * Constructor used internally by the application when making a writable
   * copy of a spreadsheet that has been read in
   *
   * @param nc the cell to copy
   */
  public Number(NumberCell nc)
  {
    super(nc);
  }

  /**
   * Sets the numerical value for this cell
   *
   * @param val the value
   */
  public void setValue(double val)
  {
    super.setValue(val);
  }

  /**
   * Copy constructor used for deep copying
   *
   * @param col the column
   * @param row the row
   * @param n the number to copy
   */
  protected Number(int col, int row, Number n)
  {
    super(col, row, n);
  }

  /**
   * Implementation of the deep copy function
   *
   * @param col the column which the new cell will occupy
   * @param row the row which the new cell will occupy
   * @return  a copy of this cell, which can then be added to the sheet
   */
  public WritableCell copyTo(int col, int row)
  {
    return new Number(col, row, this);
  }

}
