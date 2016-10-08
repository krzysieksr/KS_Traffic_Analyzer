/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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
 * The environment specific print record
 */
public class PLSRecord extends RecordData
{
  /**
   * Constructs this object from the raw data
   *
   * @param r the raw data
   */
  public PLSRecord(Record r)
  {
    super(r);
  }

  /**
   * Gets the data
   *
   * @return the binary data
   */
  public byte[] getData()
  {
    return getRecord().getData();
  }
}
