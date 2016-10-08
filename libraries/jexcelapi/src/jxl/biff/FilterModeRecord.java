/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import jxl.common.Logger;

import jxl.read.biff.Record;

/**
 * Range information for conditional formatting
 */
public class FilterModeRecord extends WritableRecordData
{
  // The logger
  private static Logger logger = Logger.getLogger(FilterModeRecord.class);

  /**
   * The data
   */
  private byte[] data;


  /**
   * Constructor
   */
  public FilterModeRecord(Record t)
  {
    super(t);

    data = getRecord().getData();
  }

  /**
   * Retrieves the data for output to binary file
   * 
   * @return the data to be written
   */
  public byte[] getData()
  {
    return data;
  }
}

