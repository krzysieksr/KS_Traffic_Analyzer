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

import jxl.common.Logger;

import java.util.Stack;

/**
 * A cell reference in a formula
 */
abstract class BinaryOperator extends Operator implements ParsedThing
{
  // The logger
  private static final Logger logger = Logger.getLogger(BinaryOperator.class);

  /**
   * Constructor
   */
  public BinaryOperator()
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
   *
   * @param s the token stack
   */
  public void getOperands(Stack s)
  {
    ParseItem o1 = (ParseItem) s.pop();
    ParseItem o2 = (ParseItem) s.pop();

    add(o1);
    add(o2);
  }

  /**
   * Gets the string version of this binary operator
   *
   * @param buf a the string buffer
   */
  public void getString(StringBuffer buf)
  {
    ParseItem[] operands = getOperands();
    operands[1].getString(buf);
    buf.append(getSymbol());
    operands[0].getString(buf);
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
    operands[1].adjustRelativeCellReferences(colAdjust, rowAdjust);
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
    operands[1].columnInserted(sheetIndex, col, currentSheet);
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
    operands[1].columnRemoved(sheetIndex, col, currentSheet);
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
    operands[1].rowInserted(sheetIndex, row, currentSheet);
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
    operands[1].rowRemoved(sheetIndex, row, currentSheet);
    operands[0].rowRemoved(sheetIndex, row, currentSheet);
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
    byte[] data = new byte[0];

    // Get the operands in reverse order to get the RPN
    for (int i = operands.length - 1; i >= 0; i--)
    {
      byte[] opdata = operands[i].getBytes();

      // Grow the array
      byte[] newdata = new byte[data.length + opdata.length];
      System.arraycopy(data, 0, newdata, 0, data.length);
      System.arraycopy(opdata, 0, newdata, data.length, opdata.length);
      data = newdata;
    }

    // Add on the operator byte
    byte[] newdata = new byte[data.length + 1];
    System.arraycopy(data, 0, newdata, 0, data.length);
    newdata[data.length] = getToken().getCode();

    return newdata;
  }

  /**
   * Abstract method which gets the binary operator string symbol
   *
   * @return the string symbol for this token
   */
  abstract String getSymbol();

  /**
   * Abstract method which gets the token for this operator
   *
   * @return the string symbol for this token
   */
  abstract Token getToken();

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing, as operators don't have cell references
   */
  void handleImportedCellReferences()
  {
    ParseItem[] operands = getOperands();
    operands[0].handleImportedCellReferences();
    operands[1].handleImportedCellReferences();
  }

}





