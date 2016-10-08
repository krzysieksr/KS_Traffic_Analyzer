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

package jxl.write.biff;

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * The default row height for cells in the workbook
 */
class DefaultRowHeightRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The default row height
   */
  private int rowHeight;

  /**
   * Indicates whether or not the default row height has been changed
   */
  private boolean changed;
  
  /**
   * Constructor
   *
   * @param height the default row height
   * @param ch TRUE if the default value has been changed, false 
   * otherwise
   */
  public DefaultRowHeightRecord(int h, boolean ch)
  {
    super(Type.DEFAULTROWHEIGHT);
    data = new byte[4];
    rowHeight = h;
    changed = ch;
  }

  /**
   * Gets the binary data for writing to the output stream
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    if (changed)
    {
      data[0] |= 0x1;
    }

    IntegerHelper.getTwoBytes(rowHeight, data, 2);
    return data;
  }
}


