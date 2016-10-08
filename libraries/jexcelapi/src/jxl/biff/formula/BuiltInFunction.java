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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;

/**
 * A built in function in a formula
 */
class BuiltInFunction extends Operator implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(BuiltInFunction.class);

  /**
   * The function
   */
  private Function function;

  /**
   * The workbook settings
   */
  private WorkbookSettings settings;

  /**
   * Constructor
   * @param ws the workbook settings
   */
  public BuiltInFunction(WorkbookSettings ws)
  {
    settings = ws;
  }

  /**
   * Constructor used when parsing a formula from a string
   *
   * @param f the function
   * @param ws the workbook settings
   */
  public BuiltInFunction(Function f, WorkbookSettings ws)
  {
    function = f;
    settings = ws;
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
    int index = IntegerHelper.getInt(data[pos], data[pos + 1]);
    function = Function.getFunction(index);
    Assert.verify(function != Function.UNKNOWN, "function code " + index);
    return 2;
  }

  /**
   * Gets the operands for this operator from the stack
   *
   * @param s the token stack
   */
  public void getOperands(Stack s)
  {
    // parameters are in the correct order, god damn them
    ParseItem[] items = new ParseItem[function.getNumArgs()];
    // modified in 2.4.3
    for (int i = function.getNumArgs() - 1; i >= 0; i--)
    {
      ParseItem pi = (ParseItem) s.pop();

      items[i] = pi;
    }

    for (int i = 0; i < function.getNumArgs(); i++)
    {
      add(items[i]);
    }
  }

  /**
   * Gets the string for this functions
   *
   * @param buf the buffer to append
   */
  public void getString(StringBuffer buf)
  {
    buf.append(function.getName(settings));
    buf.append('(');

    int numArgs = function.getNumArgs();

    if (numArgs > 0)
    {
      ParseItem[] operands = getOperands();

      // arguments are in the same order they were specified
      operands[0].getString(buf);

      for (int i = 1; i < numArgs; i++)
      {
        buf.append(',');
        operands[i].getString(buf);
      }
    }

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

    for (int i = 0; i < operands.length; i++)
    {
      operands[i].adjustRelativeCellReferences(colAdjust, rowAdjust);
    }
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
    for (int i = 0; i < operands.length; i++)
    {
      operands[i].columnInserted(sheetIndex, col, currentSheet);
    }
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
    for (int i = 0; i < operands.length; i++)
    {
      operands[i].columnRemoved(sheetIndex, col, currentSheet);
    }
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
    for (int i = 0; i < operands.length; i++)
    {
      operands[i].rowInserted(sheetIndex, row, currentSheet);
    }
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
    for (int i = 0; i < operands.length; i++)
    {
      operands[i].rowRemoved(sheetIndex, row, currentSheet);
    }
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing, as operators don't have cell references
   */
  void handleImportedCellReferences()
  {
    ParseItem[] operands = getOperands();
    for (int i = 0 ; i < operands.length ; i++)
    {
      operands[i].handleImportedCellReferences();
    }
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

    for (int i = 0; i < operands.length; i++)
    {
      byte[] opdata = operands[i].getBytes();

      // Grow the array
      byte[] newdata = new byte[data.length + opdata.length];
      System.arraycopy(data, 0, newdata, 0, data.length);
      System.arraycopy(opdata, 0, newdata, data.length, opdata.length);
      data = newdata;
    }

    // Add on the operator byte
    byte[] newdata = new byte[data.length + 3];
    System.arraycopy(data, 0, newdata, 0, data.length);
    newdata[data.length] = !useAlternateCode() ? Token.FUNCTION.getCode() :
                                                 Token.FUNCTION.getCode2();
    IntegerHelper.getTwoBytes(function.getCode(), newdata, data.length + 1);

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
    return 3;
  }
}

