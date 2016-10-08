/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import jxl.biff.IntegerHelper;

/**
 * Indicates that the function doesn't evaluate to a constant reference
 */
class MemArea extends SubExpression
{
  /**
   * Constructor
   */
  public MemArea()
  {
  }

  public void getString(StringBuffer buf)
  {
    ParseItem[] subExpression = getSubExpression();

    if (subExpression.length == 1)
    {
      subExpression[0].getString(buf);
    }
    else if (subExpression.length == 2)
    {
      subExpression[1].getString(buf);
      buf.append(':');
      subExpression[0].getString(buf);
    }
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
    // For mem areas, the first four bytes are not used
    setLength(IntegerHelper.getInt(data[pos+4], data[pos+5]));
    return 6;
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

