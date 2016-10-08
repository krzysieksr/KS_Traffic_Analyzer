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

package jxl;

/**
 * Represents a 3-D range of cells in a workbook.  This object is
 * returned by the method findByName in a workbook
 */
public interface Range
{
  /**
   * Gets the cell at the top left of this range
   *
   * @return the cell at the top left
   */
  public Cell getTopLeft();

  /**
   * Gets the cell at the bottom right of this range
   *
   * @return the cell at the bottom right
   */
  public Cell getBottomRight();

  /**
   * Gets the index of the first sheet in the range
   *
   * @return the index of the first sheet in the range
   */
  public int getFirstSheetIndex();

  /**
   * Gets the index of the last sheet in the range
   *
   * @return the index of the last sheet in the range
   */
  public int getLastSheetIndex();
}



