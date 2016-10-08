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

import jxl.common.Logger;

import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * An external sheet record, used to maintain integrity when formulas
 * are copied from read databases
 */
class ExternalNameRecord extends WritableRecordData
{
  /**
   * The logger
   */
  Logger logger = Logger.getLogger(ExternalNameRecord.class);

  /**
   * The name of the addin
   */
  private String name;

  /**
   * Constructor used for writable workbooks
   */
  public ExternalNameRecord(String n)
  {
    super(Type.EXTERNNAME);
    name = n;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[name.length() * 2 + 12];

    data[6] = (byte) name.length();
    data[7] = 0x1;
    StringHelper.getUnicodeBytes(name, data, 8);
    
    int pos = 8 + name.length() * 2;
    data[pos]   = 0x2;
    data[pos+1] = 0x0;
    data[pos+2] = 0x1c;
    data[pos+3] = 0x17;

    return data;
  }
}
