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

package jxl.biff.formula;

import java.util.Stack;

import jxl.biff.IntegerHelper;

/**
 * Indicates that the function doesn't evaluate to a constant reference
 */
class MemFunc extends SubExpression
{
  /**
   * Constructor
   */
  public MemFunc()
  {
  }

  public void getString(StringBuffer sb)
  {
  }

 /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing
   */
  void handleImportedCellReferences()
  {
  }
}




