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

package jxl.biff;

/**
 * Converts excel byte representations into integers
 */
public final class IntegerHelper
{
  /**
   * Private constructor disables the instantiation of this object
   */
  private IntegerHelper()
  {
  }

  /**
   * Gets an int from two bytes
   *
   * @param b2 the second byte
   * @param b1 the first byte
   * @return The integer value
   */
  public static int getInt(byte b1, byte b2)
  {
    int i1 = b1 & 0xff;
    int i2 = b2 & 0xff;
    int val = i2 << 8 | i1;
    return val;
  }

  /**
   * Gets an short from two bytes
   *
   * @param b2 the second byte
   * @param b1 the first byte
   * @return The short value
   */
  public static short getShort(byte b1, byte b2)
  {
    short i1  = (short) (b1 & 0xff);
    short i2  = (short) (b2 & 0xff);
    short val = (short) (i2 << 8 | i1);
    return val;
  }


  /**
   * Gets an int from four bytes, doing all the necessary swapping
   *
   * @param b1 a byte
   * @param b2 a byte
   * @param b3 a byte
   * @param b4 a byte
   * @return the integer value represented by the four bytes
   */
  public static int getInt(byte b1, byte b2, byte b3, byte b4)
  {
    int i1 = getInt(b1, b2);
    int i2 = getInt(b3, b4);

    int val = i2 << 16 | i1;
    return val;
  }

  /**
   * Gets a two byte array from an integer
   *
   * @param i the integer
   * @return the two bytes
   */
  public static byte[] getTwoBytes(int i)
  {
    byte[] bytes = new byte[2];

    bytes[0] = (byte) (i & 0xff);
    bytes[1] = (byte) ((i & 0xff00) >> 8);

    return bytes;
  }

  /**
   * Gets a four byte array from an integer
   *
   * @param i the integer
   * @return a four byte array
   */
  public static byte[] getFourBytes(int i)
  {
    byte[] bytes = new byte[4];

    int i1 = i & 0xffff;
    int i2 = (i & 0xffff0000) >> 16;

    getTwoBytes(i1, bytes, 0);
    getTwoBytes(i2, bytes, 2);

    return bytes;
  }


  /**
   * Converts an integer into two bytes, and places it in the array at the
   * specified position
   *
   * @param target the array to place the byte data into
   * @param pos the position at which to place the data
   * @param i the integer value to convert
   */
  public static void getTwoBytes(int i, byte[] target, int pos)
  {
    target[pos] = (byte) (i & 0xff);
    target[pos + 1] = (byte) ((i & 0xff00) >> 8);
  }

  /**
   * Converts an integer into four bytes, and places it in the array at the
   * specified position
   *
   * @param target the array which is to contain the converted data
   * @param pos the position in the array in which to place the data
   * @param i the integer to convert
   */
  public static void getFourBytes(int i, byte[] target, int pos)
  {
    byte[] bytes = getFourBytes(i);
    target[pos]     = bytes[0];
    target[pos + 1] = bytes[1];
    target[pos + 2] = bytes[2];
    target[pos + 3] = bytes[3];
  }
}
