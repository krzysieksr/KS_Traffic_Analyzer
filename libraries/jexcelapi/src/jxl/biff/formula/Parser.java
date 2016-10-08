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

/**
 * Interface used by the two different types of formula parser
 */
interface Parser
{
  /**
   * Parses the formula
   *
   * @exception FormulaException if an error occurs
   */
  public void parse() throws FormulaException;

  /**
   * Gets the string version of the formula
   *
   * @return the formula as a string
   */
  public String getFormula();

  /**
   * Gets the bytes for the formula. This takes into account any
   * token mapping necessary because of shared formulas
   *
   * @return the bytes in RPN
   */
  public byte[] getBytes();

  /**
   * Adjusts all the relative cell references in this formula by the
   * amount specified.  
   *
   * @param colAdjust
   * @param rowAdjust
   */
  public void adjustRelativeCellReferences(int colAdjust, int rowAdjust);


  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param col the column number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void columnInserted(int sheetIndex, int col, boolean currentSheet);

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was removed
   * @param col the column number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void columnRemoved(int sheetIndex, int col, boolean currentSheet);

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param row the column number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowInserted(int sheetIndex, int row, boolean currentSheet);

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was removed
   * @param row the column number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowRemoved(int sheetIndex, int row, boolean currentSheet);

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   *
   * @return TRUE if the formula is valid import, FALSE otherwise
   */
  public boolean handleImportedCellReferences();
}
