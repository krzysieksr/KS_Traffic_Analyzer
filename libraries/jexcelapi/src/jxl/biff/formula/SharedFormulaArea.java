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

package jxl.biff.formula;

import jxl.Cell;
import jxl.biff.CellReferenceHelper;
import jxl.biff.IntegerHelper;

/**
 * A cell reference in a formula
 */
class SharedFormulaArea extends Operand implements ParsedThing
{
  private int columnFirst;
  private int rowFirst;
  private int columnLast;
  private int rowLast;

  private boolean columnFirstRelative;
  private boolean rowFirstRelative;
  private boolean columnLastRelative;
  private boolean rowLastRelative;

  /**
   * The cell containing the formula.  Stored in order to determine
   * relative cell values
   */
  private Cell relativeTo;

  /**
   * Constructor
   *
   * @param the cell the formula is relative to
   */
  public SharedFormulaArea(Cell rt)
  {
    relativeTo = rt;
  }

  int getFirstColumn() 
  {
    return columnFirst;
  }
  
  int getFirstRow() 
  {
    return rowFirst;
  }

  int getLastColumn() 
  {
    return columnLast;
  }

  int getLastRow() 
  {
    return rowLast;
  }

  /** 
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos)
  {
    // Preserve signage on column and row values, because they will
    // probably be relative

    rowFirst =  IntegerHelper.getShort(data[pos], data[pos+1]);
    rowLast  = IntegerHelper.getShort(data[pos+2], data[pos+3]);

    int columnMask = IntegerHelper.getInt(data[pos+4], data[pos+5]);
    columnFirst = columnMask & 0x00ff;
    columnFirstRelative = ((columnMask & 0x4000) != 0);
    rowFirstRelative = ((columnMask & 0x8000) != 0);

    if (columnFirstRelative)
    {
      columnFirst = relativeTo.getColumn() + columnFirst;
    }

    if (rowFirstRelative)
    {
      rowFirst = relativeTo.getRow() + rowFirst;
    }

    columnMask = IntegerHelper.getInt(data[pos+6], data[pos+7]);
    columnLast = columnMask & 0x00ff;

    columnLastRelative = ((columnMask & 0x4000) != 0);
    rowLastRelative = ((columnMask & 0x8000) != 0);

    if (columnLastRelative)
    { 
      columnLast = relativeTo.getColumn() + columnLast;
    }

    if (rowLastRelative)
    {
      rowLast = relativeTo.getRow() + rowLast;
    }


    return 8;
  }

  public void getString(StringBuffer buf)
  {
    CellReferenceHelper.getCellReference(columnFirst, rowFirst, buf);
    buf.append(':');
    CellReferenceHelper.getCellReference(columnLast, rowLast, buf);
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[9];
    data[0] = Token.AREA.getCode();

    // Use absolute references for columns, so don't bother about
    // the col relative/row relative bits
    IntegerHelper.getTwoBytes(rowFirst, data, 1);
    IntegerHelper.getTwoBytes(rowLast, data, 3);
    IntegerHelper.getTwoBytes(columnFirst, data, 5);
    IntegerHelper.getTwoBytes(columnLast, data, 7);

    return data;
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing
   */
  void handleImportedCellReferences()
  {
  }

}



