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


import jxl.BooleanCell;
import jxl.CellType;
import jxl.biff.Type;
import jxl.format.CellFormat;


/**
 * A boolean cell's last calculated value
 */
public abstract class BooleanRecord extends CellValue
{
  /**
   * The boolean value of this cell.  If this cell represents an error, 
   * this will be false
   */
  private boolean value;

  /**
   * Constructor invoked by the user API
   * 
   * @param c the column
   * @param r the row
   * @param val the value
   */
  protected BooleanRecord(int c, int r, boolean val)
  {
    super(Type.BOOLERR, c, r);
    value = val;
  }

  /**
   * Overloaded constructor invoked from the API, which takes a cell
   * format
   * 
   * @param c the column
   * @param r the row
   * @param val the value
   * @param st the cell format
   */
  protected BooleanRecord(int c, int r, boolean val, CellFormat st)
  {
    super(Type.BOOLERR, c, r, st);
    value = val;
  }

  /**
   * Constructor used when copying a workbook
   * 
   * @param nc the number to copy
   */
  protected BooleanRecord(BooleanCell nc)
  {
    super(Type.BOOLERR, nc);
    value = nc.getValue();
  }

  /**
   * Copy constructor
   * 
   * @param c the column
   * @param r the row
   * @param br the record to copy
   */
  protected BooleanRecord(int c, int r, BooleanRecord br)
  {
    super(Type.BOOLERR, c, r, br);
    value = br.value;
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
   * Sets the value
   * 
   * @param val the boolean value
   */
  protected void setValue(boolean val)
  {
    value = val;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] celldata = super.getData();
    byte[] data = new byte[celldata.length + 2];
    System.arraycopy(celldata, 0, data, 0, celldata.length);

    if (value)
    {
      data[celldata.length] = 1;
    }

    return data;
  }

}






