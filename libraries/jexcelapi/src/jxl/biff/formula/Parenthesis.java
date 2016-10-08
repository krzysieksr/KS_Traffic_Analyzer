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

/**
 * A cell reference in a formula
 */
class Parenthesis extends Operator implements ParsedThing
{
  /** 
   * Constructor
   */
  public Parenthesis()
  {
  }

  /** 
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos)
  {
    return 0;
  }

  /** 
   * Gets the operands for this operator from the stack
   */
  public void getOperands(Stack s)
  {
    ParseItem pi = (ParseItem) s.pop();

    add(pi);
  }

  public void getString(StringBuffer buf)
  {
    ParseItem[] operands = getOperands();
    buf.append('(');
    operands[0].getString(buf);
    buf.append(')');
  }

  /**
   * Adjusts all the relative cell references in this formula by the
   * amount specified.  Used when copying formulas
   *
   * @param colAdjust the amount to add on to each relative cell reference
   * @param rowAdjust the amount to add on to each relative row reference
   */
  public void adjustRelativeCellReferences(int colAdjust, int rowAdjust)
  {
    ParseItem[] operands = getOperands();
    operands[0].adjustRelativeCellReferences(colAdjust, rowAdjust);
  }

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
  void columnInserted(int sheetIndex, int col, boolean currentSheet)
  {
    ParseItem[] operands = getOperands();
    operands[0].columnInserted(sheetIndex, col, currentSheet);
  }

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
  void columnRemoved(int sheetIndex, int col, boolean currentSheet)
  {
    ParseItem[] operands = getOperands();
    operands[0].columnRemoved(sheetIndex, col, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was inserted
   * @param row the row number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowInserted(int sheetIndex, int row, boolean currentSheet)
  {
    ParseItem[] operands = getOperands();
    operands[0].rowInserted(sheetIndex, row, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the row was removed
   * @param row the row number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void rowRemoved(int sheetIndex, int row, boolean currentSheet)
  {
    ParseItem[] operands = getOperands();
    operands[0].rowRemoved(sheetIndex, row, currentSheet);
  }

 /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing, as operators don't have cell references
   */
  void handleImportedCellReferences()
  {
    ParseItem[] operands = getOperands();
    operands[0].handleImportedCellReferences();
  }

  /**
   * Abstract method which gets the token for this operator
   *
   * @return the string symbol for this token
   */
  Token getToken()
  {
    return Token.PARENTHESIS;
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    // Get the data for the operands
    ParseItem[] operands = getOperands();
    byte[] data = operands[0].getBytes();

    // Add on the operator byte
    byte[] newdata = new byte[data.length + 1];
    System.arraycopy(data, 0, newdata, 0, data.length);
    newdata[data.length] = getToken().getCode();

    return newdata;
  }

  /**
   * Gets the precedence for this operator.  Operator precedents run from 
   * 1 to 5, one being the highest, 5 being the lowest
   *
   * @return the operator precedence
   */
  int getPrecedence()
  {
    return 4;
  }
}
