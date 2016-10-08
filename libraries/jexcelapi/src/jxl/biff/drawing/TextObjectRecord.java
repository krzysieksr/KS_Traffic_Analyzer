/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

package jxl.biff.drawing;

import jxl.common.Logger;

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.read.biff.Record;

/**
 * A TextObject (TXO) record which contains the information for comments
 */
public class TextObjectRecord extends WritableRecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(TextObjectRecord.class);

  /**
   * The raw drawing data which was read in
   */
  private byte[] data;

  /**
   * The text
   */
  private int textLength;

  /**
   * Constructor invoked when writing out this object
   *
   * @param t the text string
   */
  TextObjectRecord(String t)
  {
    super(Type.TXO);

    textLength = t.length();
  }

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   */
  public TextObjectRecord(Record t)
  {
    super(t);
    data = getRecord().getData();
    textLength = IntegerHelper.getInt(data[10], data[11]);
  }

  /**
   * Constructor
   *
   * @param d the drawing data
   */
  public TextObjectRecord(byte[] d)
  {
    super(Type.TXO);
    data = d;
  }

  /**
   * Gets the text length.  Used to determine if there is any data held
   * in the following continue records
   *
   * @return the length of the text
   */
  public int getTextLength()
  {
    return textLength;
  }

  /**
   * Expose the protected function to the SheetImpl in this package
   *
   * @return the raw record data
   */
  public byte[] getData()
  {
    if (data != null)
    {
      return data;
    }

    data = new byte[18];

    // the options
    int options = 0;
    options |= (0x1 << 1); // horizontal alignment - left
    options |= (0x1 << 4); // vertical alignment - top
    options |= (0x1 << 9); // lock text
    IntegerHelper.getTwoBytes(options, data, 0);

    // the rotation
    // no rotation

    // Length of text
    IntegerHelper.getTwoBytes(textLength, data, 10);

    // Length of formatting runs
    IntegerHelper.getTwoBytes(0x10, data, 12);

    return data;
  }
}




