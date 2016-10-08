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

package jxl.read.biff;

import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * Contains the default column width for cells in this sheet
 */
class DefaultColumnWidthRecord extends RecordData
{
  /**
   * The default columns width, in characters
   */
  private int width;

  /**
   * Constructs the def col width from the raw data
   *
   * @param t the raw data
   */
  public DefaultColumnWidthRecord(Record t)
  {
    super(t);
    byte[] data = t.getData();

    width = IntegerHelper.getInt(data[0], data[1]);
  }


  /**
   * Accessor for the default width
   *
   * @return the width
   */
  public int getWidth()
  {
    return width;
  }
}







