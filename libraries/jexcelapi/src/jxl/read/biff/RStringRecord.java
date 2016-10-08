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
import jxl.WorkbookSettings;
import jxl.biff.FormattingRecords;
import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;

/**
 * A label which is stored in the cell
 */
class RStringRecord extends CellValue implements LabelCell
{
  /**
   * The length of the label in characters
   */
  private int length;
  /**
   * The label
   */
  private String string;

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static Biff7 biff7 = new Biff7();

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param fr the formatting records
   * @param si the sheet
   * @param ws the workbook settings
   * @param dummy dummy overload to indicate a biff 7 workbook
   */
  public RStringRecord(Record t, FormattingRecords fr,
                       SheetImpl si, WorkbookSettings ws, Biff7 dummy)
  {
    super(t, fr, si);
    byte[] data = getRecord().getData();
    length = IntegerHelper.getInt(data[6], data[7]);

    string = StringHelper.getString(data, length, 8, ws);
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
   * Gets the cell contents as a string
   *
   * @return the label
   */
  public String getContents()
  {
    return string;
  }

  /**
   * Accessor for the cell type
   *
   * @return the cell type
   */
  public CellType getType()
  {
    return CellType.LABEL;
  }
}
