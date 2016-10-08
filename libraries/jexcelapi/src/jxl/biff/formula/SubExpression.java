/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import jxl.biff.IntegerHelper;

/**
 * Base class for those tokens which encapsulate a subexpression
 */
abstract class SubExpression extends Operand implements ParsedThing
{
  /**
   * The number of bytes in the subexpression
   */
  private int length;

  /**
   * The sub expression
   */
  private ParseItem[] subExpression;

  /**
   * Constructor
   */
  protected SubExpression()
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
    length = IntegerHelper.getInt(data[pos], data[pos+1]);
    return 2;
  }

  /** 
   * Gets the operands for this operator from the stack
   */
  public void getOperands(Stack s)
  {
  }

  /**
   * Gets the token representation of this item in RPN.  The Attribute
   * token is a special case, which overrides anything useful we could do
   * in the base class
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    return null;
  }


  /**
   * Gets the precedence for this operator.  Operator precedents run from 
   * 1 to 5, one being the highest, 5 being the lowest
   *
   * @return the operator precedence
   */
  int getPrecedence()
  {
    return 5;
  }

  /**
   * Accessor for the length
   *
   * @return the length of the subexpression
   */
  public int getLength()
  {
    return length;
  }

  protected final void setLength(int l)
  {
    length = l;
  }

  public void setSubExpression(ParseItem[] pi)
  {
    subExpression = pi;
  }

  protected ParseItem[] getSubExpression()
  {
    return subExpression;
  }
}
