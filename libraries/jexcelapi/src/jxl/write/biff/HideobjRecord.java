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
 * Stores options selected in the Options dialog box
 *
 * =2 if the Hide All option is turned on
 * =1 if the Show Placeholders option is turned on
 * =0 if the Show All option is turned on

 */
class HideobjRecord extends WritableRecordData
{
  /**
   * Hide object mode
   */
  private int hidemode;

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param newHideMode the hide all flag
   */
  public HideobjRecord(int newHideMode)
  {
    super(Type.HIDEOBJ);

    hidemode = newHideMode;
    data = new byte[2];

    IntegerHelper.getTwoBytes(hidemode, data, 0);
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    return data;
  }
}
