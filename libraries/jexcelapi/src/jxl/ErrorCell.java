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
 * This type represents a cell which contains an error.  This error will
 * usually, but not always be the result of some error resulting from
 * a formula
 */
public interface ErrorCell extends Cell
{
  /**
   * Gets the error code for this cell.  If this cell does not represent
   * an error, then it returns 0.  Always use the method isError() to
   * determine this prior to calling this method
   *
   * @return the error code if this cell contains an error, 0 otherwise
   */
  public int getErrorCode();
}
