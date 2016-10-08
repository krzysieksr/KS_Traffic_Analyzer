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
 * Enumeration class which contains the various alignments for data within a 
 * cell
 */
public /*final*/ class Alignment
{
  /**
   * The internal numerical repreentation of the alignment
   */
  private int value;

  /**
   * The string description of this alignment
   */
  private String string;

  /**
   * The list of alignments
   */
  private static Alignment[] alignments  = new Alignment[0];

  /**
   * Private constructor
   * 
   * @param val 
   * @param string
   */
  protected Alignment(int val,String s)
  {
    value = val;
    string = s;

    Alignment[] oldaligns = alignments;
    alignments = new Alignment[oldaligns.length + 1];
    System.arraycopy(oldaligns, 0, alignments, 0, oldaligns.length);
    alignments[oldaligns.length] = this;
  }

  /**
   * Gets the value of this alignment.  This is the value that is written to 
   * the generated Excel file
   * 
   * @return the binary value
   */
  public int getValue()
  {
    return value;
  }

  /**
   * Gets the string description of this alignment
   *
   * @return the string description
   */
  public String getDescription()
  {
    return string;
  }

  /**
   * Gets the alignment from the value
   *
   * @param val 
   * @return the alignment with that value
   */
  public static Alignment getAlignment(int val)
  {
    for (int i = 0 ; i < alignments.length ; i++)
    {
      if (alignments[i].getValue() == val)
      {
        return alignments[i];
      }
    }

    return GENERAL;
  }
  
  /**
   * The standard alignment
   */
  public static Alignment GENERAL = new Alignment(0, "general");
  /**
   * Data cells with this alignment will appear at the left hand edge of the 
   * cell
   */
  public static Alignment LEFT    = new Alignment(1, "left");
  /**
   * Data in cells with this alignment will be centred
   */
  public static Alignment CENTRE  = new Alignment(2, "centre");
  /**
   * Data in cells with this alignment will be right aligned
   */
  public static Alignment RIGHT   = new Alignment(3, "right");  
  /**
   * Data in cells with this alignment will fill the cell
   */
  public static Alignment FILL    = new Alignment(4,"fill");
  /**
   * Data in cells with this alignment will be justified
   */
  public static Alignment JUSTIFY = new Alignment(5, "justify");
}

