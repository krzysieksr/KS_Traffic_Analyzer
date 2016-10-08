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

/**
 * Interface to abstract away an in-memory output or a temporary file
 * output.  Used by the File object
 */
interface ExcelDataOutput
{
  /**
   * Appends the bytes to the end of the output
   *
   * @param d the data to write to the end of the array
   */
  public void write(byte[] bytes) throws IOException;

  /**
   * Gets the current position within the file
   *
   * @return the position within the file
   */
  public int getPosition() throws IOException;

  /**
   * Sets the data at the specified position to the contents of the array
   * 
   * @param pos the position to alter
   * @param newdata the data to modify
   */
  public void setData(byte[] newdata, int pos) throws IOException;

  /** 
   * Writes the data to the output stream
   */
  public void writeData(OutputStream out) throws IOException;

  /**
   * Called when the final compound file has been written
   */
  public void close() throws IOException;
}
