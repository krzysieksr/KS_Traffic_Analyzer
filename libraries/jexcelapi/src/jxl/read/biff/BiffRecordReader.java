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

package jxl.read.biff;

/**
 * Serves up Record objects from a biff file.  This object is used by the
 * demo programs BiffDump and ... only and has no influence whatsoever on
 * the JExcelApi reading and writing of excel sheets
 */
public class BiffRecordReader
{
  /**
   * The biff file
   */
  private File file;

  /**
   * The current record retrieved
   */
  private Record record;

  /**
   * Constructor
   *
   * @param f the biff file
   */
  public BiffRecordReader(File f)
  {
    file = f;
  }

  /**
   * Sees if there are any more records to read
   *
   * @return TRUE if there are more records, FALSE otherwise
   */
  public boolean hasNext()
  {
    return file.hasNext();
  }

  /**
   * Gets the next record
   *
   * @return the next record
   */
  public Record next()
  {
    record = file.next();
    return record;
  }

  /**
   * Gets the position of the current record in the biff file
   *
   * @return the position
   */
  public int getPos()
  {
    return file.getPos() - record.getLength() - 4;
  }
}
