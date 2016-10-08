/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the7 Free Software Foundation; either
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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.WorkbookSettings;
import jxl.biff.WorkbookMethods;

/**
 * Parses the formula passed in (either as parsed strings or as a string)
 * into a tree of operators and operands
 */
public class FormulaParser
{
  /**
   * The logger
   */
  private static final Logger logger = Logger.getLogger(FormulaParser.class);

  /**
   * The formula parser.  The object implementing this interface will either
   * parse tokens or strings
   */
  private Parser parser;

  /**
   * Constructor which creates the parse tree out of tokens
   *
   * @param tokens the list of parsed tokens
   * @param rt the cell containing the formula
   * @param es a handle to the external sheet
   * @param nt a handle to the name table
   * @param ws the workbook settings
   * @param pc the parse context
   * @exception FormulaException
   */
  public FormulaParser(byte[] tokens,
                       Cell rt,
                       ExternalSheet es,
                       WorkbookMethods nt,
                       WorkbookSettings ws)
   throws FormulaException
  {
    // A null workbook bof means that it is a writable workbook and therefore
    // must be biff8
    if (es.getWorkbookBof() != null &&
        !es.getWorkbookBof().isBiff8())
    {
      throw new FormulaException(FormulaException.BIFF8_SUPPORTED);
    }
    Assert.verify(nt != null);
    parser = new TokenFormulaParser(tokens, rt, es, nt, ws, 
                                    ParseContext.DEFAULT);
  }

  /**
   * Constructor which creates the parse tree out of tokens
   *
   * @param tokens the list of parsed tokens
   * @param rt the cell containing the formula
   * @param es a handle to the external sheet
   * @param nt a handle to the name table
   * @param ws the workbook settings
   * @param pc the parse context
   * @exception FormulaException
   */
  public FormulaParser(byte[] tokens,
                       Cell rt,
                       ExternalSheet es,
                       WorkbookMethods nt,
                       WorkbookSettings ws,
                       ParseContext pc)
   throws FormulaException
  {
    // A null workbook bof means that it is a writable workbook and therefore
    // must be biff8
    if (es.getWorkbookBof() != null &&
        !es.getWorkbookBof().isBiff8())
    {
      throw new FormulaException(FormulaException.BIFF8_SUPPORTED);
    }
    Assert.verify(nt != null);
    parser = new TokenFormulaParser(tokens, rt, es, nt, ws, pc);
  }

  /**
   * Constructor which creates the parse tree out of the string
   *
   * @param form the formula string
   * @param es the external sheet handle
   * @param nt the name table
   * @param ws the workbook settings
   */
  public FormulaParser(String form,
                       ExternalSheet es,
                       WorkbookMethods nt,
                       WorkbookSettings ws)
  {
    parser = new StringFormulaParser(form, es, nt, ws,
                                     ParseContext.DEFAULT);
  }

  /**
   * Constructor which creates the parse tree out of the string
   *
   * @param form the formula string
   * @param es the external sheet handle
   * @param nt the name table
   * @param ws the workbook settings
   * @param pc the context of the parse
   */
  public FormulaParser(String form,
                       ExternalSheet es,
                       WorkbookMethods nt,
                       WorkbookSettings ws,
                       ParseContext pc)
  {
    parser = new StringFormulaParser(form, es, nt, ws, pc);
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
    parser.adjustRelativeCellReferences(colAdjust, rowAdjust);
  }

  /**
   * Parses the formula into a parse tree
   *
   * @exception FormulaException
   */
  public void parse() throws FormulaException
  {
    parser.parse();
  }

  /**
   * Gets the formula as a string
   *
   * @return the formula as a string
   * @exception FormulaException
   */
  public String getFormula() throws FormulaException
  {
    return parser.getFormula();
  }

  /**
   * Gets the bytes for the formula. This takes into account any
   * token mapping necessary because of shared formulas
   *
   * @return the bytes in RPN
   */
  public byte[] getBytes()
  {
    return parser.getBytes();
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
  public void columnInserted(int sheetIndex, int col, boolean currentSheet)
  {
    parser.columnInserted(sheetIndex, col, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param col the column number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void columnRemoved(int sheetIndex, int col, boolean currentSheet)
  {
    parser.columnRemoved(sheetIndex, col, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param row the row number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowInserted(int sheetIndex, int row, boolean currentSheet)
  {
    parser.rowInserted(sheetIndex, row, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param row the row number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowRemoved(int sheetIndex, int row, boolean currentSheet)
  {
    parser.rowRemoved(sheetIndex, row, currentSheet);
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   *
   * @return TRUE if the formula is valid import, FALSE otherwise
   */
  public boolean handleImportedCellReferences()
  {
    return parser.handleImportedCellReferences();
  }
}
