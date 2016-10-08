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

package jxl;

import jxl.write.WritableWorkbook;
/**
 * Exposes some cell reference helper methods to the public interface.
 * This class merely delegates to the internally used reference helper
 */
public final class CellReferenceHelper
{
  /**
   * Hide the default constructor
   */
  private CellReferenceHelper()
  {
  }

  /**
   * Appends the cell reference for the column and row passed in to the string
   * buffer
   *
   * @param column the column
   * @param row the row
   * @param buf the string buffer to append
   */
  public static void getCellReference(int column, int row, StringBuffer buf)
  {
    jxl.biff.CellReferenceHelper.getCellReference(column, row, buf);
  }

  /**
   * Overloaded method which prepends $ for absolute reference
   *
   * @param column the column number
   * @param colabs TRUE if the column reference is absolute
   * @param row the row number
   * @param rowabs TRUE if the row reference is absolute
   * @param buf the string buffer
   */
  public static void getCellReference(int column,
                                      boolean colabs,
                                      int row,
                                      boolean rowabs,
                                      StringBuffer buf)
  {
    jxl.biff.CellReferenceHelper.getCellReference(column, colabs,
                                                  row, rowabs,
                                                  buf);
  }


  /**
   * Gets the cell reference for the specified column and row
   *
   * @param column the column
   * @param row the row
   * @return the cell reference
   */
  public static String getCellReference(int column, int row)
  {
    return jxl.biff.CellReferenceHelper.getCellReference(column, row);
  }

  /**
   * Gets the columnn number of the string cell reference
   *
   * @param s the string to parse
   * @return the column portion of the cell reference
   */
  public static int getColumn(String s)
  {
    return jxl.biff.CellReferenceHelper.getColumn(s);
  }

  /**
   * Gets the column letter corresponding to the 0-based column number
   *
   * @param c the column number
   * @return the letter for that column number
   */
  public static String getColumnReference(int c)
  {
    return jxl.biff.CellReferenceHelper.getColumnReference(c);
  }

  /**
   * Gets the row number of the cell reference
   * @param s the cell reference
   * @return the row number
   */
  public static int getRow(String s)
  {
    return jxl.biff.CellReferenceHelper.getRow(s);
  }

  /**
   * Sees if the column component is relative or not
   *
   * @param s the cell
   * @return TRUE if the column is relative, FALSE otherwise
   */
  public static boolean isColumnRelative(String s)
  {
    return jxl.biff.CellReferenceHelper.isColumnRelative(s);
  }

  /**
   * Sees if the row component is relative or not
   *
   * @param s the cell
   * @return TRUE if the row is relative, FALSE otherwise
   */
  public static boolean isRowRelative(String s)
  {
    return jxl.biff.CellReferenceHelper.isRowRelative(s);
  }

  /**
   * Gets the fully qualified cell reference given the column, row
   * external sheet reference etc
   *
   * @param sheet the sheet index
   * @param column the column index
   * @param row the row index
   * @param workbook the workbook
   * @param buf a string buffer
   */
  public static void getCellReference
    (int sheet, int column, int row,
     Workbook workbook, StringBuffer buf)
  {
    jxl.biff.CellReferenceHelper.getCellReference
      (sheet, column, row, (jxl.biff.formula.ExternalSheet) workbook, buf);
  }

  /**
   * Gets the fully qualified cell reference given the column, row
   * external sheet reference etc
   *
   * @param sheet the sheet
   * @param column the column
   * @param row the row
   * @param workbook the workbook
   * @param buf the buffer
   */
  public static void getCellReference(int sheet,
                                      int column,
                                      int row,
                                      WritableWorkbook workbook,
                                      StringBuffer buf)
  {
    jxl.biff.CellReferenceHelper.getCellReference
      (sheet, column, row, (jxl.biff.formula.ExternalSheet) workbook, buf);
  }

  /**
   * Gets the fully qualified cell reference given the column, row
   * external sheet reference etc
   *
   * @param sheet the sheet
   * @param column the column
   * @param colabs TRUE if the column is an absolute reference
   * @param row the row
   * @param rowabs TRUE if the row is an absolute reference
   * @param workbook the workbook
   * @param buf the string buffer
   */
  public static void getCellReference (int sheet,
                                       int column,
                                       boolean colabs,
                                       int row,
                                       boolean rowabs,
                                       Workbook workbook,
                                       StringBuffer buf)
  {
    jxl.biff.CellReferenceHelper.getCellReference
      (sheet, column, colabs, row, rowabs,
       (jxl.biff.formula.ExternalSheet) workbook, buf);
  }

  /**
   * Gets the fully qualified cell reference given the column, row
   * external sheet reference etc
   *
   * @param sheet the sheet
   * @param column the column
   * @param row the row
   * @param workbook the workbook
   * @return the cell reference in the form 'Sheet 1'!A1
   */
  public static String getCellReference (int sheet,
                                         int column,
                                         int row,
                                         Workbook workbook)
  {
    return jxl.biff.CellReferenceHelper.getCellReference
      (sheet, column, row, (jxl.biff.formula.ExternalSheet) workbook);
  }

  /**
   * Gets the fully qualified cell reference given the column, row
   * external sheet reference etc
   *
   * @param sheet the sheet
   * @param column the column
   * @param row the row
   * @param workbook the workbook
   * @return the cell reference in the form 'Sheet 1'!A1
   */
  public static String getCellReference(int sheet,
                                        int column,
                                        int row,
                                        WritableWorkbook workbook)
  {
    return jxl.biff.CellReferenceHelper.getCellReference
      (sheet, column, row, (jxl.biff.formula.ExternalSheet) workbook);
  }


  /**
   * Gets the sheet name from the cell reference string
   *
   * @param ref the cell reference
   * @return the sheet name
   */
  public static String getSheet(String ref)
  {
    return jxl.biff.CellReferenceHelper.getSheet(ref);
  }

  /**
   * Gets the cell reference for the cell
   * 
   * @param the cell
   */
  public static String getCellReference(Cell c)
  {
    return getCellReference(c.getColumn(), c.getRow());
  }

  /**
   * Gets the cell reference for the cell
   * 
   * @param c the cell
   * @param sb string buffer
   */
  public static void getCellReference(Cell c, StringBuffer sb)
  {
    getCellReference(c.getColumn(), c.getRow(), sb);
  }

}
