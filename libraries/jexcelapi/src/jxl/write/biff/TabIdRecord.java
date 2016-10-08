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
 * Contains an array of sheet tab index numbers
 */
class TabIdRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   *
   * @param sheets the number of sheets
   */
  public TabIdRecord(int sheets)
  {
    super(Type.TABID);

    data = new byte[sheets * 2];

    for (int i = 0 ; i < sheets; i++)
    {
      IntegerHelper.getTwoBytes(i+1, data, i * 2);
    }
  }

  /**
   * Gets the data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    return data;
  }
}
