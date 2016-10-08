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
 * Contains workbook level windowing attributes
 */
class Window1Record extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * The selected sheet
   */
  private int selectedSheet;

  /**
   * Constructor
   */
  public Window1Record(int selSheet)
  {
    super(Type.WINDOW1);
    
    selectedSheet = selSheet;

    // hard code the data in for now
    data = new byte[]
    {(byte) 0x68,
     (byte) 0x1,
     (byte) 0xe,
     (byte) 0x1,
     (byte) 0x5c,
     (byte) 0x3a,
     (byte) 0xbe,
     (byte) 0x23,
     (byte) 0x38,
     (byte) 0,
     (byte) 0,
     (byte) 0,
     (byte) 0,
     (byte) 0,
     (byte) 0x1,
     (byte) 0,
     (byte) 0x58,
     (byte) 0x2 };

    IntegerHelper.getTwoBytes(selectedSheet, data, 10);
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
