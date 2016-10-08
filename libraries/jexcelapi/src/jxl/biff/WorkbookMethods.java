/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

package jxl.biff;

import jxl.Sheet;
/**
 * An interface containing some common workbook methods.  This so that
 * objects which are re-used for both readable and writable workbooks
 * can still make the same method calls on a workbook
 */
public interface WorkbookMethods
{
  /**
   * Gets the specified sheet within this workbook
   *
   * @param index the zero based index of the required sheet
   * @return The sheet specified by the index
   */
  public Sheet getReadSheet(int index);

  /**
   * Gets the name at the specified index
   *
   * @param index the index into the name table
   * @return the name of the cell
   * @exception NameRangeException
   */
  public String getName(int index) throws NameRangeException;

  /**
   * Gets the index of the name record for the name
   *
   * @param name the name
   * @return the index in the name table
   */
  public int getNameIndex(String name);
}
