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
import jxl.biff.WritableRecordData;

/**
 * Stores a flag which indicates whether the file is a double stream
 * file.  For files written by JExcelAPI, this FALSE
 */
class DSFRecord extends WritableRecordData
{
  /**
   * The data to be written to the binary file
   */
  private byte[] data;

  /**
   * Constructor
   */
  public DSFRecord()
  {
    super(Type.DSF);

    // Hard code the fact that this is most assuredly not a double
    // stream file
    data = new byte[] {0,0};
  }

  /**
   * The binary data to be written out
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    return data;
  }
}
