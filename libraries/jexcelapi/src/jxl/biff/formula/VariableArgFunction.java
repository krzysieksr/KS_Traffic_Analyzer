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

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;

/**
 * A built in function in a formula.  These functions take a variable
 * number of arguments, such as a range (eg. SUM etc)
 */
class VariableArgFunction extends Operator implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(VariableArgFunction.class);

  /**
   * The function
   */
  private Function function;

  /**
   * The number of arguments
   */
  private int arguments;

  /**
   * Flag which indicates whether this was initialized from the client
   * api or from an excel sheet
   */
  private boolean readFromSheet;

  /**
   * The workbooks settings
   */
  private WorkbookSettings settings;

  /** 
   * Constructor
   */
  public VariableArgFunction(WorkbookSettings ws)
  {
    readFromSheet = true;
    settings = ws;
  }

  /**
   * Constructor used when parsing a function from a string
   *
   * @param f the function
   * @param a the number of arguments
   */
  public VariableArgFunction(Function f, int a, WorkbookSettings ws)
  {
    function = f;
    arguments = a;
    readFromSheet = false;
    settings = ws;
  }

  /** 
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   * @exception FormulaException
   */
  public int read(byte[] data, int pos) throws FormulaException
  {
    arguments = data[pos];
    int index = IntegerHelper.getInt(data[pos+1], data[pos+2]);
    function = Function.getFunction(index);

    if (function == Function.UNKNOWN)
    {
      throw new FormulaException(FormulaException.UNRECOGNIZED_FUNCTION,
                                 index);
    }

    return 3;
  }

  /** 
   * Gets the operands for this operator from the stack
   */
  public void getOperands(Stack s)
  {
    // parameters are in the correct order, god damn them
    ParseItem[] items = new ParseItem[arguments];

    for (int i = arguments - 1; i >= 0 ; i--)
    {
      ParseItem pi = (ParseItem) s.pop();
      
      items[i] = pi;
    }

    for (int i = 0 ; i < arguments; i++)
    {
      add(items[i]);
    }
  }

  public void getString(StringBuffer buf)
  {
    buf.append(function.getName(settings));
    buf.append('(');
   
    if (arguments > 0)
    {
      ParseItem[] operands = getOperands();
      if (readFromSheet)
      {
        // arguments are in the same order they were specified
        operands[0].getString(buf);
        
        for (int i = 1; i < arguments; i++)
        {
          buf.append(',');
          operands[i].getString(buf);
        }
      }
      else
      {
        // arguments are stored in the reverse order to which they
        // were specified, so iterate through them backwards
        operands[arguments - 1].getString(buf);
        
        for (int i = arguments - 2; i >= 0 ; i--)
        {
          buf.append(',');
          operands[i].getString(buf);
        }
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

    for (int i = 0 ; i < operands.length ; i++)
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
    for (int i = 0 ; i < operands.length ; i++)
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
    for (int i = 0 ; i < operands.length ; i++)
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
    for (int i = 0 ; i < operands.length ; i++)
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
    for (int i = 0 ; i < operands.length ; i++)
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
   * Gets the underlying function
   */
  Function getFunction()
  {
    return function;
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    handleSpecialCases();

    // Get the data for the operands - in the correct order
    ParseItem[] operands = getOperands();
    byte[] data = new byte[0];

    for (int i = 0 ; i < operands.length ; i++)
    {
      byte[] opdata = operands[i].getBytes();

      // Grow the array
      byte[] newdata = new byte[data.length + opdata.length];
      System.arraycopy(data, 0, newdata, 0, data.length);
      System.arraycopy(opdata, 0, newdata, data.length, opdata.length);
      data = newdata;
    }

    // Add on the operator byte
    byte[] newdata = new byte[data.length + 4];
    System.arraycopy(data, 0, newdata, 0, data.length);
    newdata[data.length] = !useAlternateCode() ? 
      Token.FUNCTIONVARARG.getCode() : Token.FUNCTIONVARARG.getCode2() ;
    newdata[data.length+1] = (byte) arguments;
    IntegerHelper.getTwoBytes(function.getCode(), newdata, data.length+2);

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

  /**
   * Handles functions which form a special case
   */
  private void handleSpecialCases()
  {
    // Handle the array functions.  Tell all the Area records to
    // use their alternative token code
    if (function == Function.SUMPRODUCT)
    {
      // Get the data for the operands - in reverse order
      ParseItem[] operands = getOperands();
      
      for (int i = operands.length - 1 ; i >= 0 ; i--)
      {
        if (operands[i] instanceof Area)
        {
          operands[i].setAlternateCode();
        }
      }
    }
  }
}
 

