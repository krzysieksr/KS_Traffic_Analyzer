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
import jxl.biff.Type;

/**
 * A password record
 */
class PasswordRecord extends RecordData
{
  /**
   * The password
   */
  private String password;
  /**
   * The binary data
   */
  private int passwordHash;

  /**
   * Constructor
   *
   * @param t the raw bytes
   */
  public PasswordRecord(Record t)
  {
    super(Type.PASSWORD);

    byte[] data = t.getData();
    passwordHash = IntegerHelper.getInt(data[0], data[1]);
  }

  /**
   * Gets the binary data for output to file
   *
   * @return the password hash
   */
  public int getPasswordHash()
  {
    return passwordHash;
  }
}
