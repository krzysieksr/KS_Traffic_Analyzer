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
import jxl.biff.IntegerHelper;
import jxl.biff.WritableRecordData;

/**
 * Index into the cell rows in an worksheet
 */
class IndexRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The numbe of rows served by this index record
   */
  private int rows;
  /**
   * The position of the BOF record in the excel output stream
   */
  private int bofPosition;
  /**
   * The number of blocks needed to hold all the rows
   */
  private int blocks;

  /**
   * The position of the current 'pointer' within the byte array
   */
  private int dataPos;
  
  /**
   * Constructor
   * 
   * @param pos the position of the BOF record
   * @param bl the number of blocks
   * @param r the number of rows
   */
  public IndexRecord(int pos, int r, int bl)
  {
    super(Type.INDEX);
    bofPosition = pos;
    rows = r;
    blocks = bl;

    // Allocate the amount of bytes required to hold all the blocks
    data = new byte[16 + 4 * blocks];
    dataPos = 16;
  }

  /**
   * Gets the binary data for output.  This writes out an empty data block, and
   * the information is filled in later on when the information becomes
   * available
   * 
   * @return the binary data
   */
  protected byte[] getData()
  {
    IntegerHelper.getFourBytes(rows, data, 8);
    return data;
  }

  /**
   * Adds another index record into the array
   * 
   * @param pos the position in the output file
   */
  void addBlockPosition(int pos)
  {
    IntegerHelper.getFourBytes(pos - bofPosition, data, dataPos);
    dataPos += 4;
  }

  /**
   * Sets the position of the data start.  This happens to be the position
   * of the DEFCOLWIDTH record
   */
  void setDataStartPosition(int pos)
  {
    IntegerHelper.getFourBytes(pos - bofPosition, data, 12);
  }
}
