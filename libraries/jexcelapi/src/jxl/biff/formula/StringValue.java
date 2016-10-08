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

package jxl.biff.formula;

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;


/**
 * A string constant operand in a formula
 */
class StringValue extends Operand implements ParsedThing
{
  /**
   * The logger
   */
  private final static Logger logger = Logger.getLogger(StringValue.class);

  /**
   * The string value
   */
  private String value;

  /**
   * The workbook settings
   */
  private WorkbookSettings settings;

  /**
   * Constructor
   */
  public StringValue(WorkbookSettings ws)
  {
    settings = ws;
  }

  /**
   * Constructor used when parsing a string formula
   *
   * @param s the string token, including quote marks
   */
  public StringValue(String s)
  {
    // remove the quotes
    value = s;
  }

  /** 
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos)
  {
    int length = data[pos] & 0xff;
    int consumed = 2;

    if ((data[pos+1] & 0x01) == 0)
    {
      value = StringHelper.getString(data, length, pos+2, settings);
      consumed += length;
    }
    else
    {
      value = StringHelper.getUnicodeString(data, length, pos+2);
      consumed += length * 2;
    }

    return consumed;
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[value.length() * 2 + 3];
    data[0] = Token.STRING.getCode();
    data[1] = (byte) (value.length());
    data[2] = 0x01;
    StringHelper.getUnicodeBytes(value, data, 3);

    return data;
  }

  /**
   * Abstract method implementation to get the string equivalent of this
   * token
   * 
   * @param buf the string to append to
   */
  public void getString(StringBuffer buf)
  {
    buf.append("\"");
    buf.append(value);
    buf.append("\"");
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing
   */
  void handleImportedCellReferences()
  {
  }

}

