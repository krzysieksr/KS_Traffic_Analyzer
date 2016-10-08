/*********************************************************************
*
*      Copyright (C) 2005 Andrew Khan
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

package jxl.biff;

/** 
 * A growable array of bytes
 */
public class ByteArray
{
  /**
   * The array grow size
   */
  private int growSize;

  /**
   * The current array
   */
  private byte[] bytes;

  /**
   * The current position
   */
  private int pos;

  // The default grow size
  private final static int defaultGrowSize = 1024;

  /**
   * Constructor
   */
  public ByteArray()
  {
    this(defaultGrowSize);
  }

  /**
   * Constructor
   *
   * @param gs
   */
  public ByteArray(int gs)
  {
    growSize = gs;
    bytes = new byte[defaultGrowSize];
    pos = 0;
  }

  /**
   * Adds a byte onto the array
   *
   * @param b the byte
   */
  public void add(byte b)
  {
    checkSize(1);
    bytes[pos] = b;
    pos++;
  }
  
  /**
   * Adds an array of bytes onto the array
   *
   * @param b the array of bytes
   */
  public void add(byte[] b)
  {
    checkSize(b.length);
    System.arraycopy(b, 0, bytes, pos, b.length);
    pos += b.length;
  }

  /**
   * Gets the complete array
   *
   * @return the array
   */
  public byte[] getBytes()
  {
    byte[] returnArray = new byte[pos];
    System.arraycopy(bytes, 0, returnArray, 0, pos);
    return returnArray;
  }

  /**
   * Checks to see if there is sufficient space left on the array.  If not,
   * then it grows the array
   *
   * @param sz the amount of bytes to add
   */
  private void checkSize(int sz)
  {
    while (pos + sz >= bytes.length)
    {
      //  Grow the array
      byte[] newArray = new byte[bytes.length + growSize];
      System.arraycopy(bytes, 0, newArray, 0, pos);
      bytes = newArray;
    }
  }
}
