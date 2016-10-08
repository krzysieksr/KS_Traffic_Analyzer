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

import jxl.common.Logger;

import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * Contains an array of RK numbers
 */
class MulRKRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(MulRKRecord.class);

  /**
   * The row  containing these numbers
   */
  private int row;
  /**
   * The first column these rk number occur on
   */
  private int colFirst;
  /**
   * The last column these rk numbers occur on
   */
  private int colLast;
  /**
   * The number of rk numbers contained in this record
   */
  private int numrks;
  /**
   * The array of rk numbers
   */
  private int[] rknumbers;
  /**
   * The array of xf indices
   */
  private int[] xfIndices;

  /**
   * Constructs the rk numbers from the raw data
   *
   * @param t the raw data
   */
  public MulRKRecord(Record t)
  {
    super(t);
    byte[] data = getRecord().getData();
    int length = getRecord().getLength();
    row = IntegerHelper.getInt(data[0], data[1]);
    colFirst = IntegerHelper.getInt(data[2], data[3]);
    colLast = IntegerHelper.getInt(data[length - 2], data[length - 1]);
    numrks = colLast - colFirst + 1;
    rknumbers = new int[numrks];
    xfIndices = new int[numrks];

    readRks(data);
  }

  /**
   * Reads the rks from the raw data
   *
   * @param data the raw data
   */
  private void readRks(byte[] data)
  {
    int pos = 4;
    int rk;
    for (int i = 0; i < numrks; i++)
    {
      xfIndices[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
      rk = IntegerHelper.getInt
        (data[pos + 2], data[pos + 3], data[pos + 4], data[pos + 5]);
      rknumbers[i] = rk;
      pos += 6;
    }
  }

  /**
   * Accessor for the row
   *
   * @return the row of containing these rk numbers
   */
  public int getRow()
  {
    return row;
  }

  /**
   * The first column containing the rk numbers
   *
   * @return the first column
   */
  public int getFirstColumn()
  {
    return colFirst;
  }

  /**
   * Accessor for the number of rk values
   *
   * @return the number of rk values
   */
  public int getNumberOfColumns()
  {
    return numrks;
  }

  /**
   * Returns a specific rk number
   *
   * @param index the rk number to return
   * @return the rk number in bits
   */
  public int getRKNumber(int index)
  {
    return rknumbers[index];
  }

  /**
   * Return a specific formatting index
   *
   * @param index the index of the cell in this group
   * @return the xf index
   */
  public int getXFIndex(int index)
  {
    return xfIndices[index];
  }
}




