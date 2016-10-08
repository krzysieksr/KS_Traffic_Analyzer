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
 * Enumeration type which describes the vertical alignment of data within a cell
 */
public /*final*/ class VerticalAlignment
{
  /**
   * The internal binary value which gets written to the generated Excel file
   */
  private int value;

  /**
   * The textual description
   */
  private String string;

  /**
   * The list of alignments
   */
  private static VerticalAlignment[] alignments  = new VerticalAlignment[0];

  /**
   * Constructor
   * 
   * @param val 
   */
  protected VerticalAlignment(int val, String s)
  {
    value = val;
    string = s;

    VerticalAlignment[] oldaligns = alignments;
    alignments = new VerticalAlignment[oldaligns.length + 1];
    System.arraycopy(oldaligns, 0, alignments, 0, oldaligns.length);
    alignments[oldaligns.length] = this;
  }

  /**
   * Accessor for the binary value
   * 
   * @return the internal binary value
   */
  public int getValue()
  {
    return value;
  }

  /**
   * Gets the textual description
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
  public static VerticalAlignment getAlignment(int val)
  {
    for (int i = 0 ; i < alignments.length ; i++)
    {
      if (alignments[i].getValue() == val)
      {
        return alignments[i];
      }
    }

    return BOTTOM;
  }


  /**
   * Cells with this specified vertical alignment will have their data
   * aligned at the top
   */
  public static VerticalAlignment TOP     = new VerticalAlignment(0, "top");
  /**
   * Cells with this specified vertical alignment will have their data
   * aligned centrally
   */
  public static VerticalAlignment CENTRE  = new VerticalAlignment(1, "centre");
  /**
   * Cells with this specified vertical alignment will have their data
   * aligned at the bottom
   */
  public static VerticalAlignment BOTTOM  = new VerticalAlignment(2, "bottom");  
  /**
   * Cells with this specified vertical alignment will have their data
   * justified
   */
  public static VerticalAlignment JUSTIFY = new VerticalAlignment(3, "Justify");
}

