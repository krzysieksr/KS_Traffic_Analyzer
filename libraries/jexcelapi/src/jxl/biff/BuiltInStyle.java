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
 * Represents a built in, rather than a user defined, style.
 * This class is used by the FormattingRecords class when writing out the hard*
 * coded styles
 */
class BuiltInStyle extends WritableRecordData
{
  /**
   * The XF index of this style
   */
  private int xfIndex;
  /**
   * The reference number of this style
   */
  private int styleNumber;

  /**
   * Constructor
   *
   * @param xfind the xf index of this style
   * @param sn the style number of this style
   */
  public BuiltInStyle(int xfind, int sn)
  {
    super(Type.STYLE);

    xfIndex = xfind;
    styleNumber = sn;
  }

  /**
   * Abstract method implementation to get the raw byte data ready to write out
   *
   * @return The byte data
   */
  public byte[] getData()
  {
    byte[] data = new byte[4];

    IntegerHelper.getTwoBytes(xfIndex, data, 0);

    // Set the built in bit
    data[1] |= 0x80;

    data[2] = (byte) styleNumber;

    // Set the outline level
    data[3] = (byte) 0xff;

    return data;
  }
}
