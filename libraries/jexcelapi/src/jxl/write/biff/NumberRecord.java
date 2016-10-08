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

package jxl.write.biff;

import java.text.DecimalFormat;
import java.text.NumberFormat;

import jxl.CellType;
import jxl.NumberCell;
import jxl.biff.DoubleHelper;
import jxl.biff.Type;
import jxl.biff.XFRecord;
import jxl.format.CellFormat;

/**
 * The record which contains numerical values.  All values are stored
 * as 64bit IEEE floating point values
 */
public abstract class NumberRecord extends CellValue
{
  /**
   * The number
   */
  private double value;

  /**
   * The java equivalent of the excel format
   */
  private NumberFormat format;

  /**
   * The formatter to convert the value into a string
   */
  private static DecimalFormat defaultFormat = new DecimalFormat("#.###");

  /**
   * Constructor invoked by the user API
   * 
   * @param c the column
   * @param r the row
   * @param val the value
   */
  protected NumberRecord(int c, int r, double val)
  {
    super(Type.NUMBER, c, r);
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
  protected NumberRecord(int c, int r, double val, CellFormat st)
  {
    super(Type.NUMBER, c, r, st);
    value = val;
  }

  /**
   * Constructor used when copying a workbook
   * 
   * @param nc the number to copy
   */
  protected NumberRecord(NumberCell nc)
  {
    super(Type.NUMBER, nc);
    value = nc.getValue();
  }

  /**
   * Copy constructor
   * 
   * @param c the column
   * @param r the row
   * @param nr the record to copy
   */
  protected NumberRecord(int c, int r, NumberRecord nr)
  {
    super(Type.NUMBER, c, r, nr);
    value = nr.value;
  }

  /**
   * Returns the content type of this cell
   * 
   * @return the content type for this cell
   */
  public CellType getType()
  {
    return CellType.NUMBER;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] celldata = super.getData();
    byte[] data = new byte[celldata.length + 8];
    System.arraycopy(celldata, 0, data, 0, celldata.length);
    DoubleHelper.getIEEEBytes(value, data, celldata.length);

    return data;
  }

  /**
   * Quick and dirty function to return the contents of this cell as a string.
   * For more complex manipulation of the contents, it is necessary to cast
   * this interface to correct subinterface
   * 
   * @return the contents of this cell as a string
   */
  public String getContents()
  {
    if (format == null)
    {
      format = ( (XFRecord) getCellFormat()).getNumberFormat();
      if (format == null)
      {
        format = defaultFormat;
      }
    }
    return format.format(value);
  }

  /**
   * Gets the double contents for this cell.
   * 
   * @return the cell contents
   */
  public double getValue()
  {
    return value;
  }

  /**
   * Sets the value of the contents for this cell
   * 
   * @param val the new value
   */
  public void setValue(double val)
  {
    value = val;
  }

  /**
   * Gets the NumberFormat used to format this cell.  This is the java 
   * equivalent of the Excel format
   *
   * @return the NumberFormat used to format the cell
   */
  public NumberFormat getNumberFormat()
  {
    return null;
  }
}



