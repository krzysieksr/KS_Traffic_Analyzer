/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

package jxl.write;

import java.util.Collection;

import jxl.CellFeatures;
import jxl.biff.BaseCellFeatures;

/**
 * Container for any additional cell features
 */
public class WritableCellFeatures extends CellFeatures
{
  // shadow the conditions in the base class so that they appear on
  // the public generated javadoc
  public static final ValidationCondition BETWEEN = BaseCellFeatures.BETWEEN;
  public static final ValidationCondition NOT_BETWEEN =
    BaseCellFeatures.NOT_BETWEEN;
  public static final ValidationCondition EQUAL = BaseCellFeatures.EQUAL;
  public static final ValidationCondition NOT_EQUAL =
    BaseCellFeatures.NOT_EQUAL;
  public static final ValidationCondition GREATER_THAN =
    BaseCellFeatures.GREATER_THAN;
  public static final ValidationCondition LESS_THAN =
    BaseCellFeatures.LESS_THAN;
  public static final ValidationCondition GREATER_EQUAL =
    BaseCellFeatures.GREATER_EQUAL;
  public static final ValidationCondition LESS_EQUAL =
    BaseCellFeatures.LESS_EQUAL;

  /**
   * Constructor
   */
  public WritableCellFeatures()
  {
    super();
  }

  /**
   * Copy constructor
   *
   * @param cf the cell to copy
   */
  public WritableCellFeatures(CellFeatures cf)
  {
    super(cf);
  }

  /**
   * Sets the cell comment
   *
   * @param s the comment
   */
  public void setComment(String s)
  {
    super.setComment(s);
  }

  /**
   * Sets the cell comment and sets the size of the text box (in cells)
   * in which the comment is displayed
   *
   * @param s the comment
   * @param width the width of the comment box in cells
   * @param height the height of the comment box in cells
   */
  public void setComment(String s, double width, double height)
  {
    super.setComment(s, width, height);
  }

  /**
   * Removes the cell comment, if present
   */
  public void removeComment()
  {
    super.removeComment();
  }


  /**
   * Removes any data validation, if present
   */
  public void removeDataValidation()
  {
    super.removeDataValidation();
  }

  /**
   * The list of items to validate for this cell.  For each object in the
   * collection, the toString() method will be called and the data entered
   * will be validated against that string
   *
   * @param c the list of valid values
   */
  public void setDataValidationList(Collection c)
  {
    super.setDataValidationList(c);
  }

  /**
   * The list  of items to validate for this cell in the form of a cell range.
   *
   * @param col1 the first column containing the data to validate against
   * @param row1 the first row containing the data to validate against
   * @param col2 the second column containing the data to validate against
   * @param row2 the second row containing the data to validate against
   */
  public void setDataValidationRange(int col1, int row1, int col2, int row2)
  {
    super.setDataValidationRange(col1, row1, col2, row2);
  }

  /**
   * Sets the data validation based upon a named range.  If the namedRange
   * is an empty string ("") then the cell is effectively made read only
   *
   * @param namedRange the workbook named range defining the validation 
   *                   boundaries
   */
  public void setDataValidationRange(String namedRange)
  {
    super.setDataValidationRange(namedRange);
  }


  /**
   * Sets the numeric value against which to validate
   *
   * @param val the number
   * @param c the validation condition
   */
  public void setNumberValidation(double val, ValidationCondition c)
  {
    super.setNumberValidation(val, c);
  }

  /**
   * Sets the numeric range against which to validate the data
   *
   * @param val1 the first number
   * @param val2 the second number
   * @param c the validation condition
   */
  public void setNumberValidation(double val1,
                                  double val2,
                                  ValidationCondition c)
  {
    super.setNumberValidation(val1, val2, c);
  }
}
