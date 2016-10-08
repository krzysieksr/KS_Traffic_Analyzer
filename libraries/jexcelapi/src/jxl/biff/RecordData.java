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

package jxl.biff;

import jxl.read.biff.Record;

/**
 * The record data within a record
 */
public abstract class RecordData
{
  /**
   * The raw data
   */
  private Record record;

  /**
   * The Biff code for this record.  This is set up when the record is
   * used for writing
   */
  private int code;

  /**
   * Constructs this object from the raw data
   *
   * @param r the raw data
   */
  protected RecordData(Record r)
  {
    record = r;
    code = r.getCode();
  }

  /**
   * Constructor used by the writable records
   *
   * @param t the type
   */
  protected RecordData(Type t)
  {
    code = t.value;
  }

  /**
   * Returns the raw data to its subclasses
   *
   * @return the raw data
   */
  protected Record getRecord()
  {
    return record;
  }

  /**
   * Accessor for the code
   *
   * @return the code
   */
  protected final int getCode()
  {
    return code;
  }
}








