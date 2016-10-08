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
 * Enumeration class for Escher record types
 */
final class EscherRecordType
{
  /**
   * The code of the item within the escher stream
   */
  private int value;

  /**
   * All escher types
   */
  private static EscherRecordType[] types = new EscherRecordType[0];

  /**
   * Constructor
   *
   * @param val the escher record value
   */
  private EscherRecordType(int val)
  {
    value = val;

    EscherRecordType[] newtypes = new EscherRecordType[types.length + 1];
    System.arraycopy(types, 0, newtypes, 0, types.length);
    newtypes[types.length] = this;
    types = newtypes;
  }

  /**
   * Accessor for the escher record value
   *
   * @return the escher record value
   */
  public int getValue()
  {
    return value;
  }

  /**
   * Accessor to get the item from a particular value
   *
   * @param val the escher record value
   * @return the type corresponding to val, or UNKNOWN if a match could not
   * be found
   */
  public static EscherRecordType getType(int val)
  {
    EscherRecordType type = UNKNOWN;

    for (int i = 0; i < types.length; i++)
    {
      if (val == types[i].value)
      {
        type = types[i];
        break;
      }
    }

    return type;
  }

  public static final EscherRecordType UNKNOWN = new EscherRecordType(0x0);
  public static final EscherRecordType DGG_CONTAINER =
    new EscherRecordType(0xf000);
  public static final EscherRecordType BSTORE_CONTAINER =
    new EscherRecordType(0xf001);
  public static final EscherRecordType DG_CONTAINER =
    new EscherRecordType(0xf002);
  public static final EscherRecordType SPGR_CONTAINER =
    new EscherRecordType(0xf003);
  public static final EscherRecordType SP_CONTAINER =
    new EscherRecordType(0xf004);

  public static final EscherRecordType DGG = new EscherRecordType(0xf006);
  public static final EscherRecordType BSE = new EscherRecordType(0xf007);
  public static final EscherRecordType DG = new EscherRecordType(0xf008);
  public static final EscherRecordType SPGR = new EscherRecordType(0xf009);
  public static final EscherRecordType SP = new EscherRecordType(0xf00a);
  public static final EscherRecordType OPT = new EscherRecordType(0xf00b);
  public static final EscherRecordType CLIENT_ANCHOR =
    new EscherRecordType(0xf010);
  public static final EscherRecordType CLIENT_DATA =
    new EscherRecordType(0xf011);
  public static final EscherRecordType CLIENT_TEXT_BOX =
    new EscherRecordType(0xf00d);
  public static final EscherRecordType SPLIT_MENU_COLORS =
    new EscherRecordType(0xf11e);
}
