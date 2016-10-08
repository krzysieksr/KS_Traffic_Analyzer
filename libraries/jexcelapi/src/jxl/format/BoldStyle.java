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
 * Enumeration class containing the various bold styles for data
 */
public /*final*/ class BoldStyle
{
  /**
   * The bold weight
   */
  private int value;

  /**
   * The description
   */
  private String string;

  /**
   * Constructor
   * 
   * @param val 
   */
  protected BoldStyle(int val, String s)
  {
    value = val;
    string = s;
  }

  /**
   * Gets the value of the bold weight.  This is the value that will be
   * written to the generated Excel file.
   * 
   * @return the bold weight
   */
  public int getValue() 
  {
    return value ; 
  }

  /**
   * Gets the string description of the bold style
   */
  public String getDescription()
  {
    return string;
  }

  /**
   * Normal style
   */
  public static final BoldStyle NORMAL  = new BoldStyle(0x190, "Normal");
  /**
   * Emboldened style
   */
  public static final BoldStyle BOLD    = new BoldStyle(0x2bc, "Bold");
}






