/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import java.io.OutputStream;
import java.io.IOException;

import jxl.common.Logger;

/**
 * Used to generate the excel biff data in memory.  This class wraps a byte
 * array
 */
class MemoryDataOutput implements ExcelDataOutput
{
  // The logger
  private static Logger logger = Logger.getLogger(MemoryDataOutput.class);

  /**
   * The excel data
   */
  private byte[] data;

  /**
   * The grow size for the array
   */
  private int growSize;

  /**
   * The current position within the array
   */
  private int pos;

  /**
   * Constructor
   */
  public MemoryDataOutput(int initialSize, int gs)
  {
    data = new byte[initialSize];
    growSize = gs;
    pos = 0;
  }

  /**
   * Writes the bytes to the end of the array, growing the array
   * as needs dictate
   *
   * @param d the data to write to the end of the array
   */
  public void write(byte[] bytes)
  {
    while (pos + bytes.length > data.length)
    {
      // Grow the array
      byte[] newdata = new byte[data.length + growSize];
      System.arraycopy(data, 0, newdata, 0, pos);
      data = newdata;
    }

    System.arraycopy(bytes, 0, data, pos, bytes.length);
    pos += bytes.length;
  }

  /**
   * Gets the current position within the file
   *
   * @return the position within the file
   */
  public int getPosition()
  {
    return pos;
  }

  /**
   * Sets the data at the specified position to the contents of the array
   * 
   * @param pos the position to alter
   * @param newdata the data to modify
   */
  public void setData(byte[] newdata, int pos)
  {
    System.arraycopy(newdata, 0, data, pos, newdata.length);
  }

  /** 
   * Writes the data to the output stream
   */
  public void writeData(OutputStream out) throws IOException
  {
    out.write(data, 0, pos);
  }

  /**
   * Called when the final compound file has been written.  No cleanup is
   * necessary for in-memory file generation
   */
  public void close() throws IOException
  {
  }
}
