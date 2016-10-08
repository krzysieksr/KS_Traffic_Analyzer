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

import jxl.common.Logger;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;

/**
 * Implementation class for the Range interface.  This merely
 * holds the raw range information, and when the time comes, it
 * interrogates the workbook for the object.
 * This does not keep handles to the objects for performance reasons,
 * as this could impact garbage collection on larger spreadsheets
 */
public class RangeImpl implements Range
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(RangeImpl.class);

  /**
   * A handle to the workbook
   */
  private WorkbookMethods workbook;

  /**
   * The sheet index containing the column at the top left
   */
  private int sheet1;

  /**
   * The column number of the cell at the top left of the range
   */
  private int column1;

  /**
   * The row number of the cell at the top left of the range
   */
  private int row1;

  /**
   * The sheet index of the cell at the bottom right
   */
  private int sheet2;

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
   * @param w the workbook
   * @param es the external sheet
   * @param s1 the sheet of the top left cell of the range
   * @param c1 the column number of the top left cell of the range
   * @param r1 the row number of the top left cell of the range
   * @param s2 the sheet of the bottom right cell
   * @param c2 the column number of the bottom right cell of the range
   * @param r2 the row number of the bottomr right cell of the range
   */
  public RangeImpl(WorkbookMethods w,
                   int s1, int c1, int r1,
                   int s2, int c2, int r2)
  {
    workbook = w;
    sheet1 = s1;
    sheet2 = s2;
    row1 = r1;
    row2 = r2;
    column1 = c1;
    column2 = c2;
  }

  /**
   * Gets the cell at the top left of this range
   *
   * @return the cell at the top left
   */
  public Cell getTopLeft()
  {
    Sheet s = workbook.getReadSheet(sheet1);

    if (column1 <  s.getColumns() &&
        row1 < s.getRows())
    {
      return s.getCell(column1, row1);
    }
    else
    {
      return new EmptyCell(column1, row1);
    }
  }

  /**
   * Gets the cell at the bottom right of this range
   *
   * @return the cell at the bottom right
   */
  public Cell getBottomRight()
  {
    Sheet s = workbook.getReadSheet(sheet2);

    if (column2 < s.getColumns() &&
        row2 < s.getRows())
    {
      return s.getCell(column2, row2);
    }
    else
    {
      return new EmptyCell(column2, row2);
    }
  }

  /**
   * Gets the index of the first sheet in the range
   *
   * @return the index of the first sheet in the range
   */
  public int getFirstSheetIndex()
  {
    return sheet1;
  }

  /**
   * Gets the index of the last sheet in the range
   *
   * @return the index of the last sheet in the range
   */
  public int getLastSheetIndex()
  {
    return sheet2;
  }
}









