/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

package jxl.biff;

import jxl.read.biff.Record;

/**
 * A continue record -  only used explicitly in special circumstances, as
 * the general continuation record is handled directly by the records
 */
public class ContinueRecord extends WritableRecordData
{
  /**
   * The data
   */
  private byte[] data;

  /**
   * Constructor
   *
   * @param t the raw bytes
   */
  public ContinueRecord(Record t)
  {
    super(t);
    data = t.getData();
  }

  /**
   * Constructor invoked when creating continue records
   *
   * @param d the data
   */
  public ContinueRecord(byte[] d)
  {
    super(Type.CONTINUE);
    data = d;
  }

  /**
   * Accessor for the binary data - used when copying
   *
   * @return the binary data
   */
  public byte[] getData()
  {
    return data;
  }

  /**
   * Accessor for the record.  Used when forcibly changing this record
   * into another type, notably a drawing record, as sometimes Excel appears
   * to switch to writing Continue records instead of MsoDrawing records
   *
   * @return the record
   */
  public Record getRecord()
  {
    return super.getRecord();
  }

}
