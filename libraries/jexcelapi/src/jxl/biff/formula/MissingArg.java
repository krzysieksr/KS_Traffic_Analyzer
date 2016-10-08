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

/**
 * Represents a missing argument in an argument list
 */
class MissingArg extends Operand implements ParsedThing
{
  /**
   * Constructor
   */
  public MissingArg()
  {
  }

  /** 
   * Reads the ptg data from the array starting at the specified position.
   * A missing argument contains no associated data
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos)
  {
    return 0;
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[1];
    data[0] = Token.MISSING_ARG.getCode();

    return data;
  }

  /**
   * Abstract method implementation to get the string equivalent of this
   * token
   * 
   * @param buf the string to append to
   */
  public void getString(StringBuffer buf)
  {
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

