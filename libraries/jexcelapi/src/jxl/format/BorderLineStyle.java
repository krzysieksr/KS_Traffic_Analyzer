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
 * The border line style
 */
public /*final*/ class BorderLineStyle
{
  /**
   * The value
   */
  private int value;

  /**
   * The string description
   */
  private String string;

  /**
   * The list of alignments
   */
  private static BorderLineStyle[] styles  = new BorderLineStyle[0];


  /**
   * Constructor
   */
  protected BorderLineStyle(int val, String s)
  {
    value = val;
    string = s;

    BorderLineStyle[] oldstyles = styles;
    styles = new BorderLineStyle[oldstyles.length + 1];
    System.arraycopy(oldstyles, 0, styles, 0, oldstyles.length);
    styles[oldstyles.length] = this;
  }

  /**
   * Gets the value for this line style
   *
   * @return the value
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
  public static BorderLineStyle getStyle(int val)
  {
    for (int i = 0 ; i < styles.length ; i++)
    {
      if (styles[i].getValue() == val)
      {
        return styles[i];
      }
    }

    return NONE;
  }
  
  public final static BorderLineStyle NONE
                       = new BorderLineStyle(0, "none");
  public final static BorderLineStyle THIN
                       = new BorderLineStyle(1, "thin");
  public final static BorderLineStyle MEDIUM
                       = new BorderLineStyle(2, "medium");
  public final static BorderLineStyle DASHED
                       = new BorderLineStyle(3, "dashed");
  public final static BorderLineStyle DOTTED
                       = new BorderLineStyle(4, "dotted");
  public final static BorderLineStyle THICK
                       = new BorderLineStyle(5, "thick");
  public final static BorderLineStyle DOUBLE
                       = new BorderLineStyle(6, "double");
  public final static BorderLineStyle HAIR
                       = new BorderLineStyle(7, "hair");
  public final static BorderLineStyle MEDIUM_DASHED 
                       = new BorderLineStyle(8, "medium dashed");
  public final static BorderLineStyle DASH_DOT     
                       = new BorderLineStyle(9, "dash dot");
  public final static BorderLineStyle MEDIUM_DASH_DOT
                       = new BorderLineStyle(0xa, "medium dash dot");
  public final static BorderLineStyle DASH_DOT_DOT
                       = new BorderLineStyle(0xb, "Dash dot dot");
  public final static BorderLineStyle MEDIUM_DASH_DOT_DOT
                       = new BorderLineStyle(0xc, "Medium dash dot dot");  
  public final static BorderLineStyle SLANTED_DASH_DOT
                       = new BorderLineStyle(0xd, "Slanted dash dot");
}
