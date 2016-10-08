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
import jxl.biff.RecordData;

/**
 * A excel9file record
 */
class Excel9FileRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Excel9FileRecord.class);

  /**
   * The template
   */
  private boolean excel9file;

  /**
   * Constructor
   *
   * @param t the record
   */
  public Excel9FileRecord(Record t)
  {
    super(t);
    excel9file = true;
  }

  /**
   * Accessor for the template mode
   *
   * @return the template mode
   */
  public boolean getExcel9File()
  {
    return excel9file;
  }
}
