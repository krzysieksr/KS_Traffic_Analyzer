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

package jxl.write;

import jxl.Cell;
import jxl.format.CellFormat;

/**
 * The interface for all writable cells
 */
public interface WritableCell extends Cell
{
  /**
   * Sets the cell format for this cell
   *
   * @param cf the cell format
   */
  public void setCellFormat(CellFormat cf);

  /**
   * A deep copy.  The returned cell still needs to be added to the sheet.
   * By not automatically adding the cell to the sheet, the client program
   * may change certain attributes, such as the value or the format
   *
   * @param col the column which the new cell will occupy
   * @param row the row which the new cell will occupy
   * @return  a copy of this cell, which can then be added to the sheet
   */
  public WritableCell copyTo(int col, int row);

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public WritableCellFeatures getWritableCellFeatures();

  /**
   * Sets the cell features
   *
   * @param cf the cell features
   */
  public void setCellFeatures(WritableCellFeatures cf);
}

