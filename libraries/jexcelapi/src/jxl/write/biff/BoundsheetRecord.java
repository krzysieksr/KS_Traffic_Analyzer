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

import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Record which stores the sheet name, the sheet type and the stream
 * position
 */
class BoundsheetRecord extends WritableRecordData
{
  /**
   * Hidden flag
   */
  private boolean hidden;

  /**
   * Chart only flag
   */
  private boolean chartOnly;

  /**
   * The name of the sheet
   */
  private String name;

  /**
   * The data to write to the output file
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param n the sheet name
   */
  public BoundsheetRecord(String n)
  {
    super(Type.BOUNDSHEET);
    name = n;
    hidden = false;
    chartOnly = false;
  }

  /**
   * Sets the hidden flag
   */
  void setHidden()
  {
    hidden = true;
  }

  /**
   * Sets the chart only flag
   */
  void setChartOnly()
  {
    chartOnly = true;
  }

  /**
   * Gets the data to write out to the binary file
   * 
   * @return the data to write out
   */
  public byte[] getData()
  {
    data = new byte[name.length() * 2 + 8];

    if (chartOnly)
    {
      data[5] = 0x02;
    }
    else
    {
      data[5] = 0; // set stream type to worksheet
    }

    if (hidden)
    {
      data[4] = 0x1;
      data[5] = 0x0;
    }
    
    data[6] = (byte) name.length();
    data[7] = 1;
    StringHelper.getUnicodeBytes(name, data, 8);

    return data;
  }
}








