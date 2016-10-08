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

package jxl.biff.drawing;

import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.read.biff.Record;

/**
 * A record which merely holds the MSODRAWINGGROUP data.  Used when copying
 * files  which contain images
 */
public class MsoDrawingGroupRecord extends WritableRecordData
{
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   */
  public MsoDrawingGroupRecord(Record t)
  {
    super(t);
    data = t.getData();
  }

  /**
   * Constructor
   *
   * @param d the data
   */
  MsoDrawingGroupRecord(byte[] d)
  {
    super(Type.MSODRAWINGGROUP);
    data = d;
  }

  /**
   * Expose the protected function to the SheetImpl in this package
   *
   * @return the raw record data
   */
  public byte[] getData()
  {
    return data;
  }
}




