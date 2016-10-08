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

/**
 * Ambiguously defined plus operator, used as a place holder when parsing
 * string formulas.  At this stage it could be either
 * a unary or binary operator - the string parser will deduce which and
 * create the appropriate type
 */
class Plus extends StringOperator
{
  /**
   * Constructor
   */
  public Plus()
  {
    super();
  }

  /**
   * Abstract method which gets the binary version of this operator
   */
  Operator getBinaryOperator()
  {
    return new Add();
  }

  /**
   * Abstract method which gets the unary version of this operator
   */
  Operator getUnaryOperator()
  {
    return new UnaryPlus();
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Does nothing, as operators don't have cell references
   */
  void handleImportedCellReferences()
  {
  }

}
