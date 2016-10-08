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
 * Places a string at the bottom of each page when the file is printed.
 * JExcelApi sets this to be blank
 */
class FooterRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The footer string
   */
  private String footer;

  /**
   * Consructor
   * 
   * @param s the footer
   */
  public FooterRecord(String s)
  {
    super(Type.FOOTER);

    footer = s;
  }

  /**
   * Consructor invoked when copying a sheets
   * 
   * @param fr the read footer record
   */
  public FooterRecord(FooterRecord fr)
  {
    super(Type.FOOTER);

    footer = fr.footer;
  }

  /**
   * Gets the binary data to write to the output file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    if (footer == null || footer.length() == 0)
    {
      data = new byte[0];
      return data;
    }

    data  = new byte[footer.length() * 2 + 3];
    IntegerHelper.getTwoBytes(footer.length(), data, 0);
    data[2] = (byte) 0x1;

    StringHelper.getUnicodeBytes(footer, data, 3);

    return data;
  }
}


