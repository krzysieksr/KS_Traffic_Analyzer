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

import jxl.Cell;
import jxl.format.CellFormat;
import jxl.write.biff.BlankRecord;

/**
 * A blank cell.  Despite not having any contents, it may contain
 * formatting information.  Such cells are typically used when creating
 * templates
 */
public class Blank extends BlankRecord implements WritableCell
{
  /**
   * Creates a cell which, when added to the sheet, will be presented at the
   * specified column and row co-ordinates
   *
   * @param c the column
   * @param r the row
   */
  public Blank(int c, int r)
  {
    super(c, r);
  }

  /**
   * Creates a cell which, when added to the sheet, will be presented at the
   * specified column and row co-ordinates
   * in the manner specified by the CellFormat parameter
   *
   * @param c the column
   * @param r the row
   * @param st the cell format
   */
  public Blank(int c, int r, CellFormat st)
  {
    super(c, r, st);
  }

  /**
   * Constructor used internally by the application when making a writable
   * copy of a spreadsheet being read in
   *
   * @param lc the cell to copy
   */
  public Blank(Cell lc)
  {
    super(lc);
  }


  /**
   * Copy constructor used for deep copying
   *
   * @param col the column
   * @param row the row
   * @param b the balnk cell to copy
   */
  protected Blank(int col, int row, Blank b)
  {
    super(col, row, b);
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
    return new Blank(col, row, this);
  }
}

