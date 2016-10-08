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

import java.util.ArrayList;
import java.util.Iterator;

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Indexes the first row record of the block and each individual cell.  
 * This is invoked by the sheets write process
 */
class DBCellRecord extends WritableRecordData
{
  /**
   * The file position of the first Row record in this block
   */
  private int rowPos;

  /**
   * The position of the start of the next cell after the first row.  This
   * is used as the offset for the cell positions
   */
  private int cellOffset;

  /**
   * The list of all cell positions in this block
   */
  private ArrayList cellRowPositions;

  /**
   * The position of this record in the file.  Vital for calculating offsets
   */
  private int  position;

  /**
   * Constructor
   * 
   * @param rp the position of this row
   */
  public DBCellRecord(int rp)
  {
    super(Type.DBCELL);
    rowPos = rp;
    cellRowPositions = new ArrayList(10);
  }

  /**
   * Sets the offset of this cell record within the sheet stream
   * 
   * @param pos the offset
   */
  void setCellOffset(int pos)
  {
    cellOffset = pos;
  }

  /**
   * Adds a cell
   * 
   * @param pos 
   */
  void addCellRowPosition(int pos)
  {
    cellRowPositions.add(new Integer(pos));
  }

  /**
   * Sets the position of this cell within the sheet stream
   * 
   * @param pos the position
   */
  void setPosition(int pos)
  {
    position = pos;
  }

  /**
   * Gets the binary data for this cell record
   * 
   * @return the binary data
   */
  protected byte[] getData()
  {
    byte[] data = new byte[4 + 2 * cellRowPositions.size()];

    // Set the offset to the first row
    IntegerHelper.getFourBytes(position - rowPos, data, 0);

    // Now add in all the cell offsets
    int pos = 4;
    int lastCellPos = cellOffset;
    Iterator i = cellRowPositions.iterator();
    while (i.hasNext())
    {
      int cellPos = ((Integer) i.next()).intValue();
      IntegerHelper.getTwoBytes(cellPos - lastCellPos, data, pos);
      lastCellPos = cellPos;
      pos += 2;
    }

    return data;
  }
}
