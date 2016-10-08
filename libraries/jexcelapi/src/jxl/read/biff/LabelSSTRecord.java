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

import jxl.CellType;
import jxl.LabelCell;
import jxl.biff.FormattingRecords;
import jxl.biff.IntegerHelper;

/**
 * A label which is stored in the shared string table
 */
class LabelSSTRecord extends CellValue implements LabelCell
{
  /**
   * The index into the shared string table
   */
  private int index;
  /**
   * The label
   */
  private String string;

  /**
   * Constructor.  Retrieves the index from the raw data and looks it up
   * in the shared string table
   *
   * @param stringTable the shared string table
   * @param t the raw data
   * @param fr the formatting records
   * @param si the sheet
   */
  public LabelSSTRecord(Record t, SSTRecord stringTable, FormattingRecords fr,
                        SheetImpl si)
  {
    super(t, fr, si);
    byte[] data = getRecord().getData();
    index = IntegerHelper.getInt(data[6], data[7], data[8], data[9]);
    string = stringTable.getString(index);
  }

  /**
   * Gets the label
   *
   * @return the label
   */
  public String getString()
  {
    return string;
  }

  /**
   * Gets this cell's contents as a string
   *
   * @return the label
   */
  public String getContents()
  {
    return string;
  }

  /**
   * Returns the cell type
   *
   * @return the cell type
   */
  public CellType getType()
  {
    return CellType.LABEL;
  }
}
