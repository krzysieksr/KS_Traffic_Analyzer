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
 * Record which specifies a print header for a work sheet
 */
class HeaderRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The print header string
   */
  private String header;

  /**
   * Constructor
   * 
   * @param s the header string
   */
  public HeaderRecord(String h)
  {
    super(Type.HEADER);

    header = h;
  }

  /**
   * Consructor invoked when copying a sheets
   * 
   * @param hr the read header record
   */
  public HeaderRecord(HeaderRecord hr)
  {
    super(Type.HEADER);

    header = hr.header;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    if (header == null || header.length() == 0)
    {
      data = new byte[0];
      return data;
    }

    data  = new byte[header.length() * 2 + 3];
    IntegerHelper.getTwoBytes(header.length(), data, 0);
    data[2] = (byte) 0x1;

    StringHelper.getUnicodeBytes(header, data, 3);

    return data;
  }
}


