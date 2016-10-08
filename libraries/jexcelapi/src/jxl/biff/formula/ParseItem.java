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

/**
 * Abstract base class for an item in a formula parse tree
 */
abstract class ParseItem
{
  // The logger
  private static Logger logger = Logger.getLogger(ParseItem.class);

  /**
   * The parent of this parse item
   */
  private ParseItem parent;

  /**
   * Volatile flag
   */
  private boolean volatileFunction;

  /**
   * Indicates that the alternative token code should be used
   * @deprecated - use the ParseContext now
   */
  private boolean alternateCode;

  /**
   * Indicates that an alternative token code should be used
   */
  private ParseContext parseContext;

  /**
   * Indicates whether this tree represents a valid formula or not.  If not
   * the parser replaces it with a valid one
   */
  private boolean valid;

  /**
   * Constructor
   */
  public ParseItem()
  {
    volatileFunction = false;
    alternateCode = false;
    valid = true;
    parseContext = parseContext.DEFAULT;
  }

  /**
   * Called by this class to initialize the parent
   */
  protected void setParent(ParseItem p)
  {
    parent = p;
  }

  /**
   * Sets the volatile flag and ripples all the way up the parse tree
   */
  protected void setVolatile()
  {
    volatileFunction = true;
    if (parent != null && !parent.isVolatile())
    {
      parent.setVolatile();
    }
  }

  /**
   * Sets the invalid flag and ripples all the way up the parse tree
   */
  protected final void setInvalid()
  {
    valid = false;
    if (parent != null)
    {
      parent.setInvalid();
    }
  }

  /**
   * Accessor for the volatile function
   *
   * @return TRUE if the formula is volatile, FALSE otherwise
   */
  final boolean isVolatile()
  {
    return volatileFunction;
  }

  /**
   * Accessor for the volatile function
   *
   * @return TRUE if the formula is volatile, FALSE otherwise
   */
  final boolean isValid()
  {
    return valid;
  }

  /**
   * Gets the string representation of this item
   * @param ws the workbook settings
   */
  abstract void getString(StringBuffer buf);

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  abstract byte[] getBytes();

  /**
   * Adjusts all the relative cell references in this formula by the
   * amount specified.  Used when copying formulas
   *
   * @param colAdjust the amount to add on to each relative cell reference
   * @param rowAdjust the amount to add on to each relative row reference
   */
  abstract void adjustRelativeCellReferences(int colAdjust, int rowAdjust);

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
  abstract void columnInserted(int sheetIndex, int col, boolean currentSheet);

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
  abstract void columnRemoved(int sheetIndex, int col, boolean currentSheet);

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
  abstract void rowInserted(int sheetIndex, int row, boolean currentSheet);

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
  abstract void rowRemoved(int sheetIndex, int row, boolean currentSheet);

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   */
  abstract void handleImportedCellReferences();

  /**
   * Tells the operands to use the alternate code
   *
   * @deprecated - use setParseContext now
   */
  protected void setAlternateCode()
  {
    alternateCode = true;
  }

  /**
   * Accessor for the alternate code flag
   *
   * @return TRUE to use the alternate code, FALSE otherwise
   * @deprecated - use setParseContext now
   */
  protected final boolean useAlternateCode()
  {
    return alternateCode;
  }

  /**
   * Tells the operands to use the alternate code
   *
   * @pc the parse context
   */
  protected void setParseContext(ParseContext pc)
  {
    parseContext = pc;
  }

  /**
   * Accessor for the alternate code flag
   *
   * @return the parse context
   */
  protected final ParseContext getParseContext()
  {
    return parseContext;
  }

}



