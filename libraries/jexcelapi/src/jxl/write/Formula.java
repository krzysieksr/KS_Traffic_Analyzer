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

import jxl.format.CellFormat;
import jxl.write.biff.FormulaRecord;

/**
 * A cell, created by user applications, which contains a numerical value
 */
public class Formula extends FormulaRecord implements WritableCell
{
  /**
   * Constructs the formula
   *
   * @param c the column
   * @param r the row
   * @param form the  formula
   */
  public Formula(int c, int r, String form)
  {
    super(c, r, form);
  }

  /**
   * Constructs a formula
   *
   * @param c the column
   * @param r the row
   * @param form the formula
   * @param st the cell style
   */
  public Formula(int c, int r, String form, CellFormat st)
  {
    super(c, r, form, st);
  }

  /**
   * Copy constructor
   *
   * @param c the column
   * @param r the row
   * @param f the record to  copy
   */
  protected Formula(int c, int r, Formula f)
  {
    super(c, r, f);
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
    return new Formula(col, row, this);
  }
}
