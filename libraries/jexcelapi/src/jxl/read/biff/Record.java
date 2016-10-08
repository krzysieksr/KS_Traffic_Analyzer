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

import java.util.ArrayList;

import jxl.common.Logger;

import jxl.biff.IntegerHelper;
import jxl.biff.Type;


/**
 * A container for the raw record data within a biff file
 */
public final class Record
{
  /**
   * The logger
   */
  private static final Logger logger = Logger.getLogger(Record.class);

  /**
   * The excel biff code
   */
  private int code;
  /**
   * The data type
   */
  private Type type;
  /**
   * The length of this record
   */
  private int length;
  /**
   * A pointer to the beginning of the actual data
   */
  private int dataPos;
  /**
   * A handle to the excel 97 file
   */
  private File file;
  /**
   * The raw data within this record
   */
  private byte[] data;

  /**
   * Any continue records
   */
  private ArrayList continueRecords;

  /**
   * Constructor
   *
   * @param offset the offset in the raw file
   * @param f the excel 97 biff file
   * @param d the data record
   */
  Record(byte[] d, int offset, File f)
  {
    code = IntegerHelper.getInt(d[offset], d[offset + 1]);
    length = IntegerHelper.getInt(d[offset + 2], d[offset + 3]);
    file = f;
    file.skip(4);
    dataPos = f.getPos();
    file.skip(length);
    type = Type.getType(code);
  }

  /**
   * Gets the biff type
   *
   * @return the biff type
   */
  public Type getType()
  {
    return type;
  }

  /**
   * Gets the length of the record
   *
   * @return the length of the record
   */
  public int getLength()
  {
    return length;
  }

  /**
   * Gets the data portion of the record
   *
   * @return the data portion of the record
   */
  public byte[] getData()
  {
    if (data == null)
    {
      data = file.read(dataPos, length);
    }

    // copy in any data from the continue records
    if (continueRecords != null)
    {
      int size = 0;
      byte[][] contData = new byte[continueRecords.size()][];
      for (int i = 0; i < continueRecords.size(); i++)
      {
        Record r = (Record) continueRecords.get(i);
        contData[i] = r.getData();
        byte[] d2 = contData[i];
        size += d2.length;
      }

      byte[] d3 = new byte[data.length + size];
      System.arraycopy(data, 0, d3, 0, data.length);
      int pos = data.length;
      for (int i = 0; i < contData.length; i++)
      {
        byte[] d2 = contData[i];
        System.arraycopy(d2, 0, d3, pos, d2.length);
        pos += d2.length;
      }

      data = d3;
    }

    return data;
  }

  /**
   * The excel 97 code
   *
   * @return the excel code
   */
  public int getCode()
  {
    return code;
  }

  /**
   * In the case of dodgy records, this method may be called to forcibly
   * set the type in order to continue processing
   *
   * @param t the forcibly overridden type
   */
  void setType(Type t)
  {
    type = t;
  }

  /**
   * Adds a continue record to this data
   *
   * @param d the continue record
   */
  public void addContinueRecord(Record d)
  {
    if (continueRecords == null)
    {
      continueRecords = new ArrayList();
    }

    continueRecords.add(d);
  }
}
