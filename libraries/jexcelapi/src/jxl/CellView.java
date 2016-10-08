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

import jxl.format.CellFormat;

/**
 * This is a bean which client applications may use to get/set various
 * properties for a row or column on a spreadsheet
 */
public final class CellView
{
  /**
   * The dimension for the associated group of cells.  For columns this
   * will be width in characters, for rows this will be the
   * height in points
   * This attribute is deprecated in favour of the size attribute
   */
  private int dimension;

  /**
   * The size for the associated group of cells.  For columns this
   * will be width in characters multiplied by 256, for rows this will be the
   * height in points
   */
  private int size;

  /**
   * Indicates whether the deprecated function was used to set the dimension
   */
  private boolean depUsed;

  /**
   * Indicates whether or not this sheet is hidden
   */
  private boolean hidden;

  /**
   * The cell format for the row/column
   */
  private CellFormat format;

  /**
   * Indicates that this column/row should be autosized
   */
  private boolean autosize;

  /**
   * Default constructor
   */
  public CellView()
  {
    hidden = false;
    depUsed = false;
    dimension = 1;
    size = 1;
    autosize = false;
  }

  /**
   * Copy constructor
   */
  public CellView(CellView cv)
  {
    hidden = cv.hidden;
    depUsed = cv.depUsed;
    dimension = cv.dimension;
    size = cv.size;
    autosize = cv.autosize;
  }

  /**
   * Sets the hidden status of this row/column
   *
   * @param h the hidden flag
   */
  public void setHidden(boolean h)
  {
    hidden = h;
  }

  /**
   * Accessor for the hidden nature of this row/column
   *
   * @return TRUE if this row/column is hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    return hidden;
  }

  /**
   * Sets the dimension for this view
   *
   * @param d the width of the column in characters, or the height of the
   *          row in 1/20ths of a point
   * @deprecated use the setSize method instead
   */
  public void setDimension(int d)
  {
    dimension = d;
    depUsed = true;
  }

  /**
   * Sets the dimension for this view
   *
   * @param d the width of the column in characters multiplied by 256,
   *          or the height of the row in 1/20ths of a point
   */
  public void setSize(int d)
  {
    size = d;
    depUsed = false;
  }

  /**
   * Gets the width of the column in characters or the height of the
   * row in 1/20ths
   *
   * @return the dimension
   * @deprecated use getSize() instead
   */
  public int getDimension()
  {
    return dimension;
  }

  /**
   * Gets the width of the column in characters multiplied by 256, or the
   * height of the row in 1/20ths of a point
   *
   * @return the dimension
   */
  public int getSize()
  {
    return size;
  }

  /**
   * Sets the cell format for this group of cells
   *
   * @param cf the format for every cell in the column/row
   */
  public void setFormat(CellFormat cf)
  {
    format = cf;
  }

  /**
   * Accessor for the cell format for this group.
   *
   * @return the format for the column/row, or NULL if no format was
   *         specified
   */
  public CellFormat getFormat()
  {
    return format;
  }

  /**
   * Accessor for the depUsed attribute
   *
   * @return TRUE if the deprecated methods were used to set the size,
   *         FALSE otherwise
   */
  public boolean depUsed()
  {
    return depUsed;
  }

  /**
   * Sets the autosize flag.  Currently, this only works for column views
   *
   * @param a autosize
   */
  public void setAutosize(boolean a)
  {
    autosize = a;
  }

  /**
   * Accessor for the autosize flag
   * NOTE:  use of the autosize function is very processor intensive, so
   * use with care
   *
   * @return TRUE if this row/column is to be autosized
   */
  public boolean isAutosize()
  {
    return autosize;
  }
}
