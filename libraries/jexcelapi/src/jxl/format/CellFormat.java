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
 * Interface for cell formats
 */
public interface CellFormat
{
  /**
   * Gets the format used by this format
   * 
   * @return the format
   */
  public Format getFormat();

  /**
   * Gets the font information used by this format
   *
   * @return the font
   */
  public Font getFont();

  /**
   * Gets whether or not the contents of this cell are wrapped
   * 
   * @return TRUE if this cell's contents are wrapped, FALSE otherwise
   */
  public boolean getWrap();

  /**
   * Gets the horizontal cell alignment
   *
   * @return the alignment
   */
  public Alignment getAlignment();

  /**
   * Gets the vertical cell alignment
   *
   * @return the alignment
   */
  public VerticalAlignment getVerticalAlignment();

  /**
   * Gets the orientation
   *
   * @return the orientation
   */
  public Orientation getOrientation();

  /**
   * Gets the line style for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public BorderLineStyle getBorder(Border border);

  /**
   * Gets the line style for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public BorderLineStyle getBorderLine(Border border);

  /**
   * Gets the colour for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   * If the specified cell does not have an associated line style, then
   * the colour the line would be is still returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public Colour getBorderColour(Border border);

  /**
   * Determines if this cell format has any borders at all.  Used to
   * set the new borders when merging a group of cells
   *
   * @return TRUE if this cell has any borders, FALSE otherwise
   */
  public boolean hasBorders();

  /**
   * Gets the background colour used by this cell
   *
   * @return the foreground colour
   */
  public Colour getBackgroundColour();

  /**
   * Gets the pattern used by this cell format
   *
   * @return the background pattern
   */
  public Pattern getPattern();

  /**
   * Gets the indentation of the cell text
   *
   * @return the indentation
   */
  public int getIndentation();

  /**
   * Gets the shrink to fit flag
   *
   * @return TRUE if this format is shrink to fit, FALSE otherise
   */
  public boolean isShrinkToFit();

  /**
   * Accessor for whether a particular cell is locked
   *
   * @return TRUE if this cell is locked, FALSE otherwise
   */
  public boolean isLocked();
}
