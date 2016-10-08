/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan, Adam Caldwell
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
import jxl.biff.Type;

/**
 * Class containing the zoom factor for display
 */
class SCLRecord extends RecordData
{
  /**
   * The numerator of the zoom
   */
  private int numerator;

  /**
   * The denominator of the zoom
   */
  private int denominator;

  /**
   * Constructs this record from the raw data
   * @param r the record
   */
  protected SCLRecord(Record r)
  {
    super(Type.SCL);

    byte[] data = r.getData();

    numerator = IntegerHelper.getInt(data[0], data[1]);
    denominator = IntegerHelper.getInt(data[2], data[3]);
  }

  /**
   * Accessor for the zoom factor
   *
   * @return the zoom factor as the nearest integer percentage
   */
  public int getZoomFactor()
  {
    return numerator * 100 / denominator;
  }
}
