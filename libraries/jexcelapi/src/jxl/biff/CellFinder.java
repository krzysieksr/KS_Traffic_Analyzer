/**********************************************************************
*
*      Copyright (C) 2008 Andrew Khan
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

import java.util.regex.Pattern;
import java.util.regex.Matcher;

import jxl.Sheet;
import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;

/**
 * Refactorisation to provide more sophisticated find cell by contents 
 * functionality
 */
public class CellFinder
{
  private Sheet sheet;

  public CellFinder(Sheet s)
  {
    sheet = s;
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   * 
   * @param contents the string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastCol the last column within the range
   * @param lastRow the last row within the range
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(String contents, 
                       int firstCol,
                       int firstRow,
                       int lastCol,  
                       int lastRow, 
                       boolean reverse)
  {
    Cell cell = null;
    boolean found = false;
    
    int numCols = lastCol - firstCol;
    int numRows = lastRow - firstRow;

    int row1 = reverse ? lastRow : firstRow;
    int row2 = reverse ? firstRow : lastRow;
    int col1 = reverse ? lastCol : firstCol;
    int col2 = reverse ? firstCol : lastCol;
    int inc = reverse ? -1 : 1;

    for (int i = 0; i <= numCols && found == false; i++)
    {
      for (int j = 0; j <= numRows && found == false; j++)
      {
        int curCol = col1 + i * inc;
        int curRow = row1 + j * inc;
        if (curCol < sheet.getColumns() && curRow <  sheet.getRows())
        {
          Cell c = sheet.getCell(curCol, curRow);
          if (c.getType() != CellType.EMPTY)
          {
            if (c.getContents().equals(contents))
            {
              cell = c;
              found = true;
            }
          }
        }
      }
    }

    return cell;
  }

  /**
   * Finds a cell within a given range of cells
   *
   * @param contents the string to match
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(String contents)
  {
    Cell cell = null;
    boolean found = false;
    
    for (int i = 0 ; i < sheet.getRows() && found == false; i++)
    {
      Cell[] row = sheet.getRow(i);
      for (int j = 0 ; j < row.length && found == false; j++)
      {
        if (row[j].getContents().equals(contents))
        {
          cell = row[j];
          found = true;
        }
      }
    }

    return cell;
  }

  /**
   * Gets the cell whose contents match the regular expressionstring passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   * 
   * @param pattern the regular expression string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastCol the last column within the range
   * @param lastRow the last row within the range
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(Pattern pattern, 
                       int firstCol,
                       int firstRow, 
                       int lastCol, 
                       int lastRow,  
                       boolean reverse)
  {
    Cell cell = null;
    boolean found = false;
    
    int numCols = lastCol - firstCol;
    int numRows = lastRow - firstRow;

    int row1 = reverse ? lastRow : firstRow;
    int row2 = reverse ? firstRow : lastRow;
    int col1 = reverse ? lastCol : firstCol;
    int col2 = reverse ? firstCol : lastCol;
    int inc = reverse ? -1 : 1;

    for (int i = 0; i <= numCols && found == false; i++)
    {
      for (int j = 0; j <= numRows && found == false; j++)
      {
        int curCol = col1 + i * inc;
        int curRow = row1 + j * inc;
        if (curCol < sheet.getColumns() && curRow <  sheet.getRows())
        {
          Cell c = sheet.getCell(curCol, curRow);
          if (c.getType() != CellType.EMPTY)
          {
            Matcher m = pattern.matcher(c.getContents());
            if (m.matches())
            {
              cell = c;
              found = true;
            }
          }
        }
      }
    }

    return cell;
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform.  This method differs
   * from the findCell methods in that only cells with labels are
   * queried - all numerical cells are ignored.  This should therefore
   * improve performance.
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  public LabelCell findLabelCell(String contents)
  {
    LabelCell cell = null;
    boolean found = false;

    for (int i = 0; i < sheet.getRows() && !found; i++)
    {
      Cell[] row = sheet.getRow(i);
      for (int j = 0; j < row.length && !found; j++)
      {
        if ((row[j].getType() == CellType.LABEL ||
             row[j].getType() == CellType.STRING_FORMULA) &&
            row[j].getContents().equals(contents))
        {
          cell = (LabelCell) row[j];
          found = true;
        }
      }
    }

    return cell;
  }
}
