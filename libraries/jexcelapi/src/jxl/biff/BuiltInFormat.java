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

import jxl.format.Format;

/**
 * The excel string for the various built in formats.  Used to present
 * the cell format information back to the user
 *
 * The difference between this class and the various format object contained
 * in the jxl.write package is that this object contains the Excel strings,
 * not their java equivalents
 */
final class BuiltInFormat implements Format, DisplayFormat
{
  /**
   * The excel format string
   */
  private String formatString;

  /**
   * The index
   */
  private int formatIndex;

  /**
   * Constructor
   *
   * @param s the format string
   * @param i the format index
   */
  private BuiltInFormat(String s, int i)
  {
    formatIndex = i;
    formatString = s;
  }

  /**
   * Accesses the excel format string which is applied to the cell
   * Note that this is the string that excel uses, and not the java
   * equivalent
   *
   * @return the cell format string
   */
  public String getFormatString()
  {
    return formatString;
  }

  /**
   * Accessor for the index style of this format
   *
   * @return the index for this format
   */
  public int getFormatIndex()
  {
    return formatIndex;
  }
  /**
   * Accessor to see whether this format has been initialized
   *
   * @return TRUE if initialized, FALSE otherwise
   */
  public boolean isInitialized()
  {
    return true;
  }
  /**
   * Initializes this format with the specified index number
   *
   * @param pos the position of this format record in the workbook
   */
  public void initialize(int pos)
  {
  }

  /**
   * Accessor to determine whether or not this format is built in
   *
   * @return TRUE if this format is a built in format, FALSE otherwise
   */
  public boolean isBuiltIn()
  {
    return true;
  }

  /**
   * Equals method
   *
   * @return TRUE if the two built in formats are equal, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof BuiltInFormat))
    {
      return false;
    }

    BuiltInFormat bif = (BuiltInFormat) o;
    return (formatIndex == bif.formatIndex);
  }

  /**
   * The list of built in formats
   */
  public static BuiltInFormat[] builtIns = new BuiltInFormat[0x32];

  // Populate the built ins
  static
  {
    builtIns[0x0] = new BuiltInFormat("", 0);
    builtIns[0x1] = new BuiltInFormat("0", 1);
    builtIns[0x2] = new BuiltInFormat("0.00", 2);
    builtIns[0x3] = new BuiltInFormat("#,##0", 3);
    builtIns[0x4] = new BuiltInFormat("#,##0.00", 4);
    builtIns[0x5] = new BuiltInFormat("($#,##0_);($#,##0)", 5);
    builtIns[0x6] = new BuiltInFormat("($#,##0_);[Red]($#,##0)", 6);
    builtIns[0x7] = new BuiltInFormat("($#,##0_);[Red]($#,##0)", 7);
    builtIns[0x8] = new BuiltInFormat("($#,##0.00_);[Red]($#,##0.00)", 8);
    builtIns[0x9] = new BuiltInFormat("0%", 9);
    builtIns[0xa] = new BuiltInFormat("0.00%", 10);
    builtIns[0xb] = new BuiltInFormat("0.00E+00", 11);
    builtIns[0xc] = new BuiltInFormat("# ?/?", 12);
    builtIns[0xd] = new BuiltInFormat("# ??/??", 13);
    builtIns[0xe] = new BuiltInFormat("dd/mm/yyyy", 14);
    builtIns[0xf] = new BuiltInFormat("d-mmm-yy", 15);
    builtIns[0x10] = new BuiltInFormat("d-mmm", 16);
    builtIns[0x11] = new BuiltInFormat("mmm-yy", 17);
    builtIns[0x12] = new BuiltInFormat("h:mm AM/PM", 18);
    builtIns[0x13] = new BuiltInFormat("h:mm:ss AM/PM", 19);
    builtIns[0x14] = new BuiltInFormat("h:mm", 20);
    builtIns[0x15] = new BuiltInFormat("h:mm:ss", 21);
    builtIns[0x16] = new BuiltInFormat("m/d/yy h:mm", 22);
    builtIns[0x25] = new BuiltInFormat("(#,##0_);(#,##0)", 0x25);
    builtIns[0x26] = new BuiltInFormat("(#,##0_);[Red](#,##0)", 0x26);
    builtIns[0x27] = new BuiltInFormat("(#,##0.00_);(#,##0.00)", 0x27);
    builtIns[0x28] = new BuiltInFormat("(#,##0.00_);[Red](#,##0.00)", 0x28);
    builtIns[0x29] = new BuiltInFormat
      ("_(*#,##0_);_(*(#,##0);_(*\"-\"_);(@_)", 0x29);
    builtIns[0x2a] = new BuiltInFormat
      ("_($*#,##0_);_($*(#,##0);_($*\"-\"_);(@_)", 0x2a);
    builtIns[0x2b] = new BuiltInFormat
      ("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);(@_)", 0x2b);
    builtIns[0x2c] = new BuiltInFormat
      ("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);(@_)", 0x2c);
    builtIns[0x2d] = new BuiltInFormat("mm:ss", 0x2d);
    builtIns[0x2e] = new BuiltInFormat("[h]mm:ss", 0x2e);
    builtIns[0x2f] = new BuiltInFormat("mm:ss.0", 0x2f);
    builtIns[0x30] = new BuiltInFormat("##0.0E+0", 0x30);
    builtIns[0x31] = new BuiltInFormat("@", 0x31);
  }
}

