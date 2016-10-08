/*********************************************************************
*
*      Copyright (C) 2009 Andrew Khan
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
 * A hideobj record
 */
class HideobjRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(HideobjRecord.class);

  /**
   * The hide obj mode
   */
  private int hidemode;

  /**
   * Constructor
   *
   * @param t the record
   */
  public HideobjRecord(Record t)
  {
    super(t);
    byte[] data = t.getData();
    hidemode = IntegerHelper.getInt(data[0], data[1]);
  }

  /**
   * Accessor for the hide mode mode
   *
   * @return the hide mode
   */
  public int getHideMode()
  {
    return hidemode;
  }
}
