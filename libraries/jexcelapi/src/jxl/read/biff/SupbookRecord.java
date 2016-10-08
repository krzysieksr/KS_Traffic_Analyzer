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

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.StringHelper;

/**
 * A record containing the references to the various sheets (internal and
 * external) referenced by formulas in this workbook
 */
public class SupbookRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(SupbookRecord.class);

  /**
   * The type of this supbook record
   */
  private Type type;

  /**
   * The number of sheets - internal & external supbooks only
   */
  private int numSheets;

  /**
   * The name of the external file
   */
  private String fileName;

  /**
   * The names of the external sheets
   */
  private String[] sheetNames;

  /**
   * The type of supbook this refers to
   */
  private static class Type {};


  public static final Type INTERNAL = new Type();
  public static final Type EXTERNAL = new Type();
  public static final Type ADDIN    = new Type();
  public static final Type LINK     = new Type();
  public static final Type UNKNOWN  = new Type();

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param ws the workbook settings
   */
  SupbookRecord(Record t, WorkbookSettings ws)
  {
    super(t);
    byte[] data = getRecord().getData();

    // First deduce the type
    if (data.length == 4)
    {
      if (data[2] == 0x01 && data[3] == 0x04)
      {
        type = INTERNAL;
      }
      else if (data[2] == 0x01 && data[3] == 0x3a)
      {
        type = ADDIN;
      }
      else
      {
        type = UNKNOWN;
      }
    }
    else if (data[0] == 0 && data[1] == 0)
    {
      type = LINK;
    }
    else
    {
      type = EXTERNAL;
    }

    if (type == INTERNAL)
    {
      numSheets = IntegerHelper.getInt(data[0], data[1]);
    }

    if (type == EXTERNAL)
    {
      readExternal(data, ws);
    }
  }

  /**
   * Reads the external data records
   *
   * @param data the data
   * @param ws the workbook settings
   */
  private void readExternal(byte[] data, WorkbookSettings ws)
  {
    numSheets = IntegerHelper.getInt(data[0], data[1]);

    // subtract file name encoding from the length
    int ln = IntegerHelper.getInt(data[2], data[3]) - 1;
    int pos = 0;

    if (data[4] == 0)
    {
      // non-unicode string
      int encoding = data[5];
      pos = 6;
      if (encoding == 0)
      {
        fileName = StringHelper.getString(data, ln, pos, ws);
        pos += ln;
      }
      else
      {
        fileName = getEncodedFilename(data, ln, pos);
        pos += ln;
      }
    }
    else
    {
      // unicode string
      int encoding = IntegerHelper.getInt(data[5], data[6]);
      pos = 7;
      if (encoding == 0)
      {
        fileName = StringHelper.getUnicodeString(data, ln, pos);
        pos += ln * 2;
      }
      else
      {
        fileName = getUnicodeEncodedFilename(data, ln, pos);
        pos += ln * 2;
      }
    }

    sheetNames = new String[numSheets];

    for (int i = 0; i < sheetNames.length; i++)
    {
      ln = IntegerHelper.getInt(data[pos], data[pos + 1]);

      if (data[pos + 2] == 0x0)
      {
        sheetNames[i] = StringHelper.getString(data, ln, pos + 3, ws);
        pos += ln + 3;
      }
      else if (data[pos + 2] == 0x1)
      {
        sheetNames[i] = StringHelper.getUnicodeString(data, ln, pos + 3);
        pos += ln * 2 + 3;
      }
    }
  }

  /**
   * Gets the type of this supbook record
   *
   * @return the type of this supbook
   */
  public Type getType()
  {
    return type;
  }

  /**
   * Gets the number of sheets.  This will only be non-zero for internal
   * and external supbooks
   *
   * @return the number of sheets
   */
  public int getNumberOfSheets()
  {
    return numSheets;
  }

  /**
   * Gets the name of the external file
   *
   * @return the name of the external file
   */
  public String getFileName()
  {
    return fileName;
  }

  /**
   * Gets the name of the external sheet
   *
   * @param i the index of the external sheet
   * @return the name of the sheet
   */
  public String getSheetName(int i)
  {
    return sheetNames[i];
  }

  /**
   * Gets the data - used when copying a spreadsheet
   *
   * @return the raw external sheet data
   */
  public byte[] getData()
  {
    return getRecord().getData();
  }

  /**
   * Gets the encoded string from the data array
   *
   * @param data the data
   * @param ln length of the string
   * @param pos the position in the array
   * @return the string
   */
  private String getEncodedFilename(byte[] data, int ln, int pos)
  {
    StringBuffer buf = new StringBuffer();
    int endpos = pos + ln;
    while (pos <  endpos)
    {
      char c = (char) data[pos];

      if (c == '\u0001')
      {
        // next character is a volume letter
        pos++;
        c = (char) data[pos];
        buf.append(c);
        buf.append(":\\\\");
      }
      else if (c == '\u0002')
      {
        // file is on the same volume
        buf.append('\\');
      }
      else if (c == '\u0003')
      {
        // down directory
        buf.append('\\');
      }
      else if (c == '\u0004')
      {
        // up directory
        buf.append("..\\");
      }
      else
      {
        // just add on the character
        buf.append(c);
      }

      pos++;
    }

    return buf.toString();
  }

  /**
   * Gets the encoded string from the data array
   *
   * @param data the data
   * @param ln length of the string
   * @param pos the position in the array
   * @return the string
   */
  private String getUnicodeEncodedFilename(byte[] data, int ln, int pos)
  {
    StringBuffer buf = new StringBuffer();
    int endpos = pos + ln * 2;
    while (pos <  endpos)
    {
      char c = (char) IntegerHelper.getInt(data[pos], data[pos + 1]);

      if (c == '\u0001')
      {
        // next character is a volume letter
        pos += 2;
        c = (char) IntegerHelper.getInt(data[pos], data[pos + 1]);
        buf.append(c);
        buf.append(":\\\\");
      }
      else if (c == '\u0002')
      {
        // file is on the same volume
        buf.append('\\');
      }
      else if (c == '\u0003')
      {
        // down directory
        buf.append('\\');
      }
      else if (c == '\u0004')
      {
        // up directory
        buf.append("..\\");
      }
      else
      {
        // just add on the character
        buf.append(c);
      }

      pos += 2;
    }

    return buf.toString();
  }
}
