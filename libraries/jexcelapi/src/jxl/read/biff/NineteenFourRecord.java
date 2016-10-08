/*********************************************************************
*
*      Copyright (C) 2001 Andrew Khan
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

import jxl.biff.RecordData;

/**
 * Identifies the date system as the 1904 system or not
 */
class NineteenFourRecord extends RecordData
{
  /**
   * The base year for dates
   */
  private boolean nineteenFour;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   */
  public NineteenFourRecord(Record t)
  {
    super(t);

    byte[] data = getRecord().getData();

    nineteenFour = data[0] == 1 ? true : false;

  }

  /**
   * Accessor to see whether this spreadsheets dates are based around
   * 1904
   *
   * @return true if this workbooks dates are based around the 1904
   *              date system
   */
  public boolean is1904()
  {
    return nineteenFour;
  }


}


