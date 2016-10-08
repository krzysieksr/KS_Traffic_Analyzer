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

import jxl.common.Assert;

import jxl.BooleanCell;
import jxl.CellType;
import jxl.biff.FormattingRecords;

/**
 * A boolean cell last calculated value
 */
class BooleanRecord extends CellValue implements BooleanCell
{
  /**
   * Indicates whether this cell contains an error or a boolean
   */
  private boolean error;

  /**
   * The boolean value of this cell.  If this cell represents an error,
   * this will be false
   */
  private boolean value;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param fr  the formatting records
   * @param si the sheet
   */
  public BooleanRecord(Record t, FormattingRecords fr, SheetImpl si)
  {
    super(t, fr, si);
    error = false;
    value = false;

    byte[] data = getRecord().getData();

    error = (data[7] == 1);

    if (!error)
    {
      value = data[6] == 1 ? true : false;
    }
  }

  /**
   * Interface method which queries whether this cell contains an error.
   * Returns TRUE if it does, otherwise returns FALSE.
   *
   * @return TRUE if this cell is an error, FALSE otherwise
   */
  public boolean isError()
  {
    return error;
  }

  /**
   * Interface method which Gets the boolean value stored in this cell.  If
   * this cell contains an error, then returns FALSE.  Always query this cell
   *  type using the accessor method isError() prior to calling this method
   *
   * @return TRUE if this cell contains TRUE, FALSE if it contains FALSE or
   * an error code
   */
  public boolean getValue()
  {
    return value;
  }

  /**
   * Returns the numerical value as a string
   *
   * @return The numerical value of the formula as a string
   */
  public String getContents()
  {
    Assert.verify(!isError());

    // return Boolean.toString(value) - only available in 1.4 or later
    return (new Boolean(value)).toString();
  }

  /**
   * Returns the cell type
   *
   * @return The cell type
   */
  public CellType getType()
  {
    return CellType.BOOLEAN;
  }

  /**
   * A special case which overrides the method in the subclass to get
   * hold of the raw data
   *
   * @return the record
   */
  public Record getRecord()
  {
    return super.getRecord();
  }
}






