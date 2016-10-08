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

package jxl.biff.drawing;

import jxl.common.Logger;

/**
 * The client data
 */
class ClientData extends EscherAtom
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(ClientData.class);

  /**
   * The raw data
   */
  private byte[] data;

  /**
   * Constructor
   *
   * @param erd the record data
   */
  public ClientData(EscherRecordData erd)
  {
    super(erd);
  }

  /**
   * Constructor
   */
  public ClientData()
  {
    super(EscherRecordType.CLIENT_DATA);
  }

  /**
   * Accessor for the raw data
   *
   * @return the binary data
   */
  byte[] getData()
  {
    data = new byte[0];
    return setHeaderData(data);
  }
}
