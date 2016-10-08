/*********************************************************************
*
*      Copyright (C) 2005 Andrew Khan
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

import jxl.common.Logger;

/**
 * An cell reference error which occurs in a formula
 */
class CellReferenceError extends Operand implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CellReferenceError.class);

  /**
   * Constructor
   */
  public CellReferenceError()
  {
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
    // the data is unused - just return the four bytes

    return 4;
  }

  /**
   * Gets the cell reference as a string for this item
   *
   * @param buf the string buffer to populate
   */
  public void getString(StringBuffer buf)
  {
    buf.append(FormulaErrorCode.REF.getDescription());
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[5];
    data[0] = Token.REFERR.getCode();

    // bytes 1-5 are unused

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
