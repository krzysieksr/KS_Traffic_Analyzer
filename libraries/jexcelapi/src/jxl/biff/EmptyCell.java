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

import jxl.CellFeatures;
import jxl.CellType;
import jxl.format.Alignment;
import jxl.format.CellFormat;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;

/**
 * An empty cell.  Represents an empty, as opposed to a blank cell
 * in the workbook
 */
public class EmptyCell implements WritableCell
{
  /**
   * The row of this empty cell
   */
  private int row;
  /**
   * The column number of this empty cell
   */
  private int col;

  /**
   * Constructs an empty cell at the specified position
   *
   * @param c the zero based column
   * @param r the zero based row
   */
  public EmptyCell(int c, int r)
  {
    row = r;
    col = c;
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
    return col;
  }

  /**
   * Returns the content type of this cell
   *
   * @return the content type for this cell
   */
  public CellType getType()
  {
    return CellType.EMPTY;
  }

  /**
   * Quick and dirty function to return the contents of this cell as a string.
   *
   * @return an empty string
   */
  public String getContents()
  {
    return "";
  }

  /**
   * Accessor for the format which is applied to this cell
   *
   * @return the format applied to this cell
   */
  public CellFormat getCellFormat()
  {
    return null;
  }

  /**
   * Dummy override
   * @param flag the hidden flag
   */
  public void setHidden(boolean flag)
  {
  }

  /**
   * Dummy override
   * @param flag dummy
   */
  public void setLocked(boolean flag)
  {
  }

  /**
   * Dummy override
   * @param align dummy
   */
  public void setAlignment(Alignment align)
  {
  }

  /**
   * Dummy override
   * @param valign dummy
   */
  public void setVerticalAlignment(VerticalAlignment valign)
  {
  }

  /**
   * Dummy override
   * @param line dummy
   * @param border dummy
   */
  public void setBorder(Border border, BorderLineStyle line)
  {
  }

  /**
   * Dummy override
   * @param cf dummy
   */
  public void setCellFormat(CellFormat cf)
  {
  }

  /**
   * Dummy override
   * @param cf dummy
   * @deprecated
   */
  public void setCellFormat(jxl.CellFormat cf)
  {
  }

  /**
   * Indicates whether or not this cell is hidden, by virtue of either
   * the entire row or column being collapsed
   *
   * @return TRUE if this cell is hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    return false;
  }

  /**
   * Implementation of the deep copy function
   *
   * @param c the column which the new cell will occupy
   * @param r the row which the new cell will occupy
   * @return  a copy of this cell, which can then be added to the sheet
   */
  public WritableCell copyTo(int c, int r)
  {
    return new EmptyCell(c, r);
  }

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public CellFeatures getCellFeatures()
  {
    return null;
  }

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public WritableCellFeatures getWritableCellFeatures()
  {
    return null;
  }

  /**
   * Accessor for the cell features
   */
  public void setCellFeatures(WritableCellFeatures wcf)
  {
  }

}


