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
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Stores the string result of a formula calculation.  This record
 * occurs immediately after the formula
 */
class StringRecord extends WritableRecordData
{
  /**
   * The string value
   */
  private String value;

  /**
   * Constructor
   */
  public StringRecord(String val)
  {
    super(Type.STRING);

    value = val;
  }

  /**
   * The binary data to be written out
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[value.length() * 2 + 3];
    IntegerHelper.getTwoBytes(value.length(), data, 0);
    data[2] = 0x01; // unicode
    StringHelper.getUnicodeBytes(value, data, 3);

    return data;
  }
}
