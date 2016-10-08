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
 * A password record.  Thanks to Michael Matthews for sending me the
 * code to actually store the password for the sheet
 */
class PasswordRecord extends WritableRecordData
{
  /**
   * The password
   */
  private String password;
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param pw the password
   */
  public PasswordRecord(String pw)
  {
    super(Type.PASSWORD);

    password = pw;

    if (pw == null)
    {
      data = new byte[2];
      IntegerHelper.getTwoBytes(0, data, 0);
    }
    else
    {
      byte [] passwordBytes = pw.getBytes();
      int passwordHash = 0;
      for (int a = 0; a < passwordBytes.length; a++)
      {
       int shifted = rotLeft15Bit(passwordBytes[a], a + 1);
       passwordHash ^= shifted;
      }
      passwordHash ^= passwordBytes.length;
      passwordHash ^= 0xCE4B;

      data = new byte[2];
      IntegerHelper.getTwoBytes(passwordHash, data, 0);
    }
  }

  /**
   * Constructor
   * 
   * @param ph the password hash code
   */
  public PasswordRecord(int ph)
  {
    super(Type.PASSWORD);

    data = new byte[2];
    IntegerHelper.getTwoBytes(ph, data, 0);
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

  /**
   * Rotate the value by 15 bits.  Thanks to Michael for this
   *
   * @param val
   * @param rotate
   * @return int
   */
  private int rotLeft15Bit(int val, int rotate)
  {
    val = val &0x7FFF;

    for(; rotate > 0; rotate--) 
    {
      if((val & 0x4000) != 0) 
      {
        val = ((val << 1) & 0x7FFF) + 1;
      }
      else
      {
        val = (val << 1) & 0x7FFF;
      }
    }

    return val;
  }
}
