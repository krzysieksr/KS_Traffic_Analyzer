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
 * The location of a border
 */
public /*final*/ class Border
{
  /**
   * The string description
   */
  private String string;

  /**
   * Constructor
   */
  protected Border(String s)
  {
    string = s;
  }

  /**
   * Gets the description
   */
  public String getDescription()
  {
    return string;
  }

  public final static Border NONE   = new Border("none");
  public final static Border ALL    = new Border("all");
  public final static Border TOP    = new Border("top");
  public final static Border BOTTOM = new Border("bottom");
  public final static Border LEFT   = new Border("left");
  public final static Border RIGHT  = new Border("right");
}

