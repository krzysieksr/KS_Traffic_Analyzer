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

import jxl.common.Logger;

import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Writes out some arbitrary record data.  Used during the debug process
 */
class ArbitraryRecord extends WritableRecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(ArbitraryRecord.class);

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   *
   * @param type the biff code
   * @param d the data
   */
  public ArbitraryRecord(int type, byte[] d)
  {
    super(Type.createType(type));

    data = d;
    logger.warn("ArbitraryRecord of type " + type + " created");
  }

  /**
   * Retrieves the data to be written to the binary file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    return data;
  }
}








