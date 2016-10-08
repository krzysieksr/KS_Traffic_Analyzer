/*********************************************************************
*
*      Copyright (C) 2009 Andrew Khan
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

import jxl.WorkbookSettings;
import jxl.biff.RecordData;
import jxl.biff.StringHelper;
import jxl.biff.Type;

/**
 * A write access record
 */
class WriteAccessRecord extends RecordData
{
  /**
   * The write access user name
   */
  private String wauser;

  /**
   * Constructor
   *
   * @param t the raw bytes
   * @param isBiff8 Is record BIFF8 (else BIFF7)
   */
  public WriteAccessRecord(Record t, boolean isBiff8, WorkbookSettings ws)
  {
    super(Type.WRITEACCESS);

    byte[] data = t.getData();
    if (isBiff8)
    {
        wauser= StringHelper.getUnicodeString(data, 112/2, 0);
    }
    else
    {
        // BIFF7 does not use unicode encoding in string
        int length= data[1];
        wauser= StringHelper.getString(data, length, 1, ws);
    }
  }

  /**
   * Gets the binary data for output to file
   *
   * @return write access user name
   */
  public String getWriteAccess()
  {
    return wauser;
  }
}
