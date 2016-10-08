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

package jxl.biff.formula;

import java.util.Stack;

import jxl.common.Assert;

/**
 * Ambiguously defined operator, used as a place holder when parsing
 * string formulas.  At this stage it could be either
 * a unary or binary operator - the string parser will deduce which and
 * create the appropriate type
 */
abstract class StringOperator extends Operator
{
  /**
   * Constructor
   */
  protected StringOperator()
  {
    super();
  }

  /** 
   * Gets the operands for this operator from the stack.  Does nothing
   * here
   */
  public void getOperands(Stack s)
  {
    Assert.verify(false);
  }

  /**
   * Gets the precedence for this operator.  Does nothing here
   *
   * @return the operator precedence
   */
  int getPrecedence()
  {
    Assert.verify(false);
    return 0;
  }

  /**
   * Gets the token representation of this item in RPN.  Does nothing here
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    Assert.verify(false);
    return null;
  }
  
  /**
   * Gets the string representation of this item
   */
  void getString(StringBuffer buf)
  {
    Assert.verify(false);
  }

  /**
   * Default behaviour is to do nothing
   *
   * @param colAdjust the amount to add on to each relative cell reference
   * @param rowAdjust the amount to add on to each relative row reference
   */
  public void adjustRelativeCellReferences(int colAdjust, int rowAdjust)
  {
    Assert.verify(false);
  }


  /**
   * Default behaviour is to do nothing
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param col the column number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  void columnInserted(int sheetIndex, int col, boolean currentSheet)
  {
    Assert.verify(false);
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
    Assert.verify(false);
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
    Assert.verify(false);
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
    Assert.verify(false);
  }

  /**
   * Abstract method which gets the binary version of this operator
   */
  abstract Operator getBinaryOperator();

  /**
   * Abstract method which gets the unary version of this operator
   */
  abstract Operator getUnaryOperator();
}
