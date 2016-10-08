/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

/**
 * Split menu colours escher record
 */
class SplitMenuColors extends EscherAtom
{
  /**
   * The binary data
   */
  private byte[] data;

  /**
   * Constructor
   *
   * @param erd escher record data
   */
  public SplitMenuColors(EscherRecordData erd)
  {
    super(erd);
  }

  /**
   * Constructor
   */
  public SplitMenuColors()
  {
    super(EscherRecordType.SPLIT_MENU_COLORS);
    setVersion(0);
    setInstance(4);

    data = new byte[]
      {(byte) 0x0d, (byte) 0x00, (byte) 0x00, (byte) 0x08,
       (byte) 0x0c, (byte) 0x00, (byte) 0x00, (byte) 0x08,
       (byte) 0x17, (byte) 0x00, (byte) 0x00, (byte) 0x08,
       (byte) 0xf7, (byte) 0x00, (byte) 0x00, (byte) 0x10};
  }

  /**
   * The binary data
   *
   * @return the binary data
   */
  byte[] getData()
  {
    return setHeaderData(data);
  }
}
