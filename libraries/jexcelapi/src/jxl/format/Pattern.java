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

package jxl.format;


/**
 * Enumeration class which contains the various patterns available within
 * the standard Excel pattern palette
 */
public /*final*/ class Pattern
{
  /**
   * The internal numerical representation of the colour
   */
  private int value;

  /**
   * The textual description
   */
  private String string;

  /**
   * The list of patterns
   */
  private static Pattern[] patterns  = new Pattern[0];

  /**
   * Private constructor
   * 
   * @param val 
   * @param s
   */
  protected Pattern(int val, String s)
  {
    value = val;
    string = s;

    Pattern[] oldcols = patterns;
    patterns = new Pattern[oldcols.length + 1];
    System.arraycopy(oldcols, 0, patterns, 0, oldcols.length);
    patterns[oldcols.length] = this;
  }

  /**
   * Gets the value of this pattern.  This is the value that is written to 
   * the generated Excel file
   * 
   * @return the binary value
   */
  public int getValue()
  {
    return value;
  }

  /**
   * Gets the textual description
   *
   * @return the string
   */
  public String getDescription()
  {
    return string;
  }

  /**
   * Gets the pattern from the value
   *
   * @param val 
   * @return the pattern with that value
   */
  public static Pattern getPattern(int val)
  {
    for (int i = 0 ; i < patterns.length ; i++)
    {
      if (patterns[i].getValue() == val)
      {
        return patterns[i];
      }
    }

    return NONE;
  }

  public final static Pattern NONE    = new Pattern(0x0, "None");
  public final static Pattern SOLID   = new Pattern(0x1, "Solid");

  public final static Pattern GRAY_50 = new Pattern(0x2, "Gray 50%");
  public final static Pattern GRAY_75 = new Pattern(0x3, "Gray 75%");
  public final static Pattern GRAY_25 = new Pattern(0x4, "Gray 25%");

  public final static Pattern PATTERN1 = new Pattern(0x5, "Pattern 1");
  public final static Pattern PATTERN2 = new Pattern(0x6, "Pattern 2");
  public final static Pattern PATTERN3 = new Pattern(0x7, "Pattern 3");
  public final static Pattern PATTERN4 = new Pattern(0x8, "Pattern 4");
  public final static Pattern PATTERN5 = new Pattern(0x9, "Pattern 5");
  public final static Pattern PATTERN6 = new Pattern(0xa, "Pattern 6");
  public final static Pattern PATTERN7 = new Pattern(0xb, "Pattern 7");
  public final static Pattern PATTERN8 = new Pattern(0xc, "Pattern 8");
  public final static Pattern PATTERN9 = new Pattern(0xd, "Pattern 9");
  public final static Pattern PATTERN10 = new Pattern(0xe, "Pattern 10");
  public final static Pattern PATTERN11 = new Pattern(0xf, "Pattern 11");
  public final static Pattern PATTERN12 = new Pattern(0x10, "Pattern 12");
  public final static Pattern PATTERN13 = new Pattern(0x11, "Pattern 13");
  public final static Pattern PATTERN14 = new Pattern(0x12, "Pattern 14");
}











