/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * The gutters record
 */
public class GuttersRecord extends RecordData
{
  private int width;
  private int height;
  private int rowOutlineLevel;
  private int columnOutlineLevel;

  /**
   * Constructs this object from the raw data
   *
   * @param r the raw data
   */
  public GuttersRecord(Record r)
  {
    super(r);
    
    byte[] data = getRecord().getData();
    width = IntegerHelper.getInt(data[0], data[1]);
    height = IntegerHelper.getInt(data[2], data[3]);
    rowOutlineLevel = IntegerHelper.getInt(data[4], data[5]);
    columnOutlineLevel = IntegerHelper.getInt(data[6], data[7]);
  }

  int getRowOutlineLevel()
  {
    return rowOutlineLevel;
  }

  int getColumnOutlineLevel()
  {
    return columnOutlineLevel;
  }

}
