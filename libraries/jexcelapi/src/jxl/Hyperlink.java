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

import java.io.File;
import java.net.URL;

/**
 * Hyperlink information.  Only URLs or file links are supported
 *
 * Hyperlinks may apply to a range of cells; in such cases the methods
 * getRow and getColumn return the cell at the top left of the range
 * the hyperlink refers to.  Hyperlinks have no specific cell format
 * information applied to them, so the getCellFormat method will return null
 */
public interface Hyperlink
{
  /**
   * Returns the row number of this cell
   *
   * @return the row number of this cell
   */
  public int getRow();

  /**
   * Returns the column number of this cell
   *
   * @return the column number of this cell
   */
  public int getColumn();

  /**
   * Gets the range of cells which activate this hyperlink
   * The get sheet index methods will all return -1, because the
   * cells will all be present on the same sheet
   *
   * @return the range of cells which activate the hyperlink
   */
  public Range getRange();

  /**
   * Determines whether this is a hyperlink to a file
   *
   * @return TRUE if this is a hyperlink to a file, FALSE otherwise
   */
  public boolean isFile();

  /**
   * Determines whether this is a hyperlink to a web resource
   *
   * @return TRUE if this is a URL
   */
  public boolean isURL();

  /**
   * Determines whether this is a hyperlink to a location in this workbook
   *
   * @return TRUE if this is a link to an internal location
   */
  public boolean isLocation();

  /**
   * Returns the row number of the bottom right cell
   *
   * @return the row number of this cell
   */
  public int getLastRow();

  /**
   * Returns the column number of the bottom right cell
   *
   * @return the column number of this cell
   */
  public int getLastColumn();

  /**
   * Gets the URL referenced by this Hyperlink
   *
   * @return the URL, or NULL if this hyperlink is not a URL
   */
  public URL getURL();

  /**
   * Returns the local file eferenced by this Hyperlink
   *
   * @return the file, or NULL if this hyperlink is not a file
   */
  public File getFile();
}

