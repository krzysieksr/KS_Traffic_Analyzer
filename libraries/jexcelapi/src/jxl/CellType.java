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

package jxl;

/**
 * An enumeration type listing the available content types for a cell
 */
public final class CellType
{

  /**
   * The text description of this cell type
   */
  private String description;

  /**
   * Private constructor
   * @param desc the description of this type
   */
  private CellType(String desc)
  {
    description = desc;
  }

  /**
   * Returns a string description of this cell
   *
   * @return the string description for this type
   */
  public String toString()
  {
    return description;
  }

  /**
   * An empty cell can still contain formatting information and comments
   */
  public static final CellType EMPTY           = new CellType("Empty");
  /**
   */
  public static final CellType LABEL           = new CellType("Label");
  /**
   */
  public static final CellType NUMBER          = new CellType("Number");
  /**
   */
  public static final CellType BOOLEAN         = new CellType("Boolean");
  /**
   */
  public static final CellType ERROR           = new CellType("Error");
  /**
   */
  public static final CellType NUMBER_FORMULA  =
                                         new CellType("Numerical Formula");
  /**
   */
  public static final CellType DATE_FORMULA  = new CellType("Date Formula");
  /**
   */
  public static final CellType STRING_FORMULA  = new CellType("String Formula");
  /**
   */
  public static final CellType BOOLEAN_FORMULA =
                                          new CellType("Boolean Formula");
  /**
   */
  public static final CellType FORMULA_ERROR   = new CellType("Formula Error");
  /**
   */
  public static final CellType DATE            = new CellType("Date");

}


