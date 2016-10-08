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
 * The default column width for a workbook
 */
class DefaultColumnWidth extends WritableRecordData
{
  /**
   * The default column width
   */
  private int width;
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param w the default column width
   */
  public DefaultColumnWidth(int w)
  {
    super(Type.DEFCOLWIDTH);
    width = w;
    data = new byte[2];
    IntegerHelper.getTwoBytes(width, data, 0);
  }

  /**
   * Gets the binary data for writing to the stream
   * 
   * @return the binary data
   */
  protected byte[] getData()
  {
    return data;
  }
}
