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

package jxl.read.biff;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.StringHelper;

/**
 * A boundsheet record, which contains the worksheet name
 */
class BoundsheetRecord extends RecordData
{
  /**
   * The offset into the sheet
   */
  private int offset;
  /**
   * The type of sheet this is
   */
  private byte typeFlag;
  /**
   * The visibility flag
   */
  private byte visibilityFlag;
  /**
   * The length of the worksheet name
   */
  private int length;
  /**
   * The worksheet name
   */
  private String name;

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static Biff7 biff7 = new Biff7();

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param s the workbook settings
   */
  public BoundsheetRecord(Record t, WorkbookSettings s)
  {
    super(t);
    byte[] data = getRecord().getData();
    offset = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);
    typeFlag = data[5];
    visibilityFlag = data[4];
    length = data[6];

    if (data[7] == 0)
    {
      // Standard ASCII encoding
      byte[] bytes = new byte[length];
      System.arraycopy(data, 8, bytes, 0, length);
      name = StringHelper.getString(bytes, length, 0, s);
    }
    else
    {
      // little endian Unicode encoding
      byte[] bytes = new byte[length * 2];
      System.arraycopy(data, 8, bytes, 0, length * 2);
      name = StringHelper.getUnicodeString(bytes, length, 0);
    }
  }


  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param biff7 a dummy value to tell the record to interpret the
   *              data as biff7
   */
  public BoundsheetRecord(Record t, Biff7 biff7)
  {
    super(t);
    byte[] data = getRecord().getData();
    offset = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);
    typeFlag = data[5];
    visibilityFlag = data[4];
    length = data[6];
    byte[] bytes = new byte[length];
    System.arraycopy(data, 7, bytes, 0, length);
    name = new String(bytes);
  }

  /**
   * Accessor for the worksheet name
   *
   * @return the worksheet name
   */
  public String getName()
  {
    return name;
  }

  /**
   * Accessor for the hidden flag
   *
   * @return TRUE if this is a hidden sheet, FALSE otherwise
   */
  public boolean isHidden()
  {
    return visibilityFlag != 0;
  }

  /**
   * Accessor to determine if this is a worksheet, or some other nefarious
   * type of object
   *
   * @return TRUE if this is a worksheet, FALSE otherwise
   */
  public boolean isSheet()
  {
    return typeFlag == 0;
  }

  /**
   * Accessor to determine if this is a chart
   *
   * @return TRUE if this is a chart, FALSE otherwise
   */
  public boolean isChart()
  {
    return typeFlag == 2;
  }

}




