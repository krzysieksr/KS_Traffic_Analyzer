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
 * Enumeration type which describes the page orientation
 */
public class PageOrder 
{
  /**
   * Constructor
   */
  private PageOrder() 
  {
  }
    
  /**
   * Top to Down then Right.
   */
  public static PageOrder DOWN_THEN_RIGHT   = new PageOrder();
  
  /**
   * Left to Right then Down.
   */
  public static PageOrder RIGHT_THEN_DOWN  = new PageOrder();
}
