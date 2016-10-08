/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

import jxl.biff.BaseCellFeatures;

/**
 * Container for any additional cell features
 */
public class CellFeatures extends BaseCellFeatures
{
  /**
   * Constructor
   */
  public CellFeatures()
  {
    super();
  }

  /**
   * Copy constructor
   *
   * @param cf cell to copy
   */
  protected CellFeatures(CellFeatures cf)
  {
    super(cf);
  }

  /**
   * Accessor for the cell comment
   *
   * @return the cell comment, or NULL if this cell doesn't have
   *         a comment associated with it
   */
  public String getComment()
  {
    return super.getComment();
  }

  /**
   * Gets the data validation list
   *
   * @return the data validation list
   */
  public String getDataValidationList()
  {
    return super.getDataValidationList();
  }

 /**
   * Gets the range of cells to which the data validation applies.  If the
   * validation applies to just this cell, this will be reflected in the 
   * returned range
   *
   * @return the range to which the same validation extends, or NULL if this
   *         cell doesn't have a validation
   */
  public Range getSharedDataValidationRange()
  {
    return super.getSharedDataValidationRange();
  }
}
