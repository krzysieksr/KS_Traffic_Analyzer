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
 * Marks the beginning of the user interface record
 */
class InterfaceHeaderRecord extends WritableRecordData
{
  /**
   * Constructor
   */
  public InterfaceHeaderRecord()
  {
    super(Type.INTERFACEHDR);
  }

  /**
   * Gets the binary data
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    // Return the character encoding
    byte[] data = new byte[] 
    { (byte) 0xb0, (byte) 0x04};
    return data; 
  }
}


