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
 * Record which contains the bounds of the sheet
 */
class DimensionRecord extends WritableRecordData
{
  /**
   * The number of rows in the sheet
   */
  private int numRows;
  /**
   * The number of columns in the sheet
   */
  private int numCols;

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   * 
   * @param c the number of columns
   * @param r the number of rows
   */
  public DimensionRecord(int r, int c)
  {
    super(Type.DIMENSION);
    numRows = r;
    numCols = c;

    data = new byte[14];

    IntegerHelper.getFourBytes(numRows, data, 4);
    IntegerHelper.getTwoBytes(numCols, data, 10);
  }

  /**
   * Gets the binary data to be written to the output file
   * 
   * @return the binary data
   */
  protected byte[] getData()
  {
    return data;
  }

}
