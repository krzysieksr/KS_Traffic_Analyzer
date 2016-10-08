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

import jxl.LabelCell;
import jxl.format.CellFormat;
import jxl.write.biff.LabelRecord;

/**
 * A cell containing text which may be created by user applications
 */
public class Label extends LabelRecord implements WritableCell, LabelCell
{
  /**
   * Creates a cell which, when added to the sheet, will be presented at the
   * specified column and row co-ordinates and will contain the specified text
   *
   * @param c the column
   * @param cont the text
   * @param r the row
   */
  public Label(int c, int r, String cont)
  {
    super(c, r, cont);
  }

  /**
   * Creates a cell which, when added to the sheet, will be presented at the
   * specified column and row co-ordinates and will present the specified text
   * in the manner specified by the CellFormat parameter
   *
   * @param c the column
   * @param cont the data
   * @param r the row
   * @param st the cell format
   */
  public Label(int c, int r, String cont, CellFormat st)
  {
    super(c, r, cont, st);
  }

  /**
   * Copy constructor used for deep copying
   *
   * @param col the column
   * @param row the row
   * @param l the label to copy
   */
  protected Label(int col, int row, Label l)
  {
    super(col, row, l);
  }

  /**
   * Constructor used internally by the application when making a writable
   * copy of a spreadsheet being read in
   *
   * @param lc the label to copy
   */
  public Label(LabelCell lc)
  {
    super(lc);
  }

  /**
   * Sets the string contents of this cell
   *
   * @param s the new data
   */
  public void setString(String s)
  {
    super.setString(s);
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
    return new Label(col, row, this);
  }
}

