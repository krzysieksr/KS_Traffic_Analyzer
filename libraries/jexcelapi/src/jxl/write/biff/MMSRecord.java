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

import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Stores the number of addmen and delmenu groups in the book stream
 */
class MMSRecord extends WritableRecordData
{
  /**
   * The number of menu items added
   */
  private byte numMenuItemsAdded;
  /**
   * The number of menu items deleted
   */
  private byte numMenuItemsDeleted;
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param menuItemsAdded the number of menu items added
   * @param menuItemsDeleted the number of menu items deleted
   */
  public MMSRecord(int menuItemsAdded, int menuItemsDeleted)
  {
    super(Type.MMS);

    numMenuItemsAdded   = (byte) menuItemsAdded;
    numMenuItemsDeleted = (byte) menuItemsDeleted;
    
    data = new byte[2];

    data[0] = numMenuItemsAdded;
    data[1] = numMenuItemsDeleted;
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
