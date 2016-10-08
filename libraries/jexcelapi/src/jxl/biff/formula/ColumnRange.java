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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.biff.CellReferenceHelper;

/**
 * A class to hold range information across two entire columns
 */
class ColumnRange extends Area
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(ColumnRange.class);

  /**
   * Constructor
   */
  ColumnRange()
  {
    super();
  }

  /**
   * Constructor invoked when parsing a string formula
   *
   * @param s the string to parse
   */
  ColumnRange(String s)
  {
    int seppos = s.indexOf(":");
    Assert.verify(seppos != -1);
    String startcell = s.substring(0, seppos);
    String endcell = s.substring(seppos + 1);

    int columnFirst = CellReferenceHelper.getColumn(startcell);
    int rowFirst = 0;
    int columnLast = CellReferenceHelper.getColumn(endcell);
    int rowLast = 0xffff;

    boolean columnFirstRelative =
      CellReferenceHelper.isColumnRelative(startcell);
    boolean rowFirstRelative = false;
    boolean columnLastRelative = CellReferenceHelper.isColumnRelative(endcell);
    boolean rowLastRelative = false;

    setRangeData(columnFirst, columnLast,
                 rowFirst, rowLast,
                 columnFirstRelative, columnLastRelative,
                 rowFirstRelative, rowLastRelative);
  }

  /**
   * Gets the string representation of this item
   *
   * @param buf the string buffer
   */
  public void getString(StringBuffer buf)
  {
    CellReferenceHelper.getColumnReference(getFirstColumn(), buf);
    buf.append(':');
    CellReferenceHelper.getColumnReference(getLastColumn(), buf);
  }
}
