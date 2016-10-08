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

import jxl.Cell;
import jxl.WorkbookSettings;
import jxl.biff.WorkbookMethods;

/**
 * Parses the excel ptgs into a parse tree
 */
class TokenFormulaParser implements Parser
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(TokenFormulaParser.class);

  /**
   * The Excel ptgs
   */
  private byte[] tokenData;

  /**
   * The cell containing the formula.  This is used in order to determine
   * relative cell values
   */
  private Cell relativeTo;

  /**
   * The current position within the array
   */
  private int pos;

  /**
   * The parse tree
   */
  private ParseItem root;

  /**
   * The hash table of items that have been parsed
   */
  private Stack tokenStack;

  /**
   * A reference to the workbook which holds the external sheet
   * information
   */
  private ExternalSheet workbook;

  /**
   * A reference to the name table
   */
  private WorkbookMethods nameTable;

  /**
   * The workbook settings
   */
  private WorkbookSettings settings;

  /**
   * The parse context
   */
  private ParseContext parseContext;

  /**
   * Constructor
   */
  public TokenFormulaParser(byte[] data, 
                            Cell c, 
                            ExternalSheet es, 
                            WorkbookMethods nt,
                            WorkbookSettings ws,
                            ParseContext pc)
  {
    tokenData = data;
    pos = 0;
    relativeTo = c;
    workbook = es;
    nameTable = nt;
    tokenStack = new Stack();
    settings = ws;
    parseContext = pc;
    
    Assert.verify(nameTable != null);
  }

  /**
   * Parses the sublist of tokens.  In most cases this will equate to
   * the full list
   *
   * @exception FormulaException
   */
  public void parse() throws FormulaException
  {
    parseSubExpression(tokenData.length);

    // Finally, there should be one thing left on the stack.  Get that
    // and add it to the root node
    root = (ParseItem) tokenStack.pop();

    Assert.verify(tokenStack.empty());

  }

  /**
   * Parses the sublist of tokens.  In most cases this will equate to
   * the full list
   *
   * @param len the length of the subexpression to parse
   * @exception FormulaException
   */
  private void parseSubExpression(int len) throws FormulaException
  {
    int tokenVal = 0;
    Token t = null;
    
    // Indicates that we are parsing the incredibly complicated and
    // hacky if construct that MS saw fit to include, the gits
    Stack ifStack = new Stack();

    // The end position of the sub-expression
    int endpos = pos + len;

    while (pos < endpos)
    {
      tokenVal = tokenData[pos];
      pos++;

      t = Token.getToken(tokenVal);

      if (t == Token.UNKNOWN)
      {
        throw new FormulaException
          (FormulaException.UNRECOGNIZED_TOKEN, tokenVal);
      }

      Assert.verify(t != Token.UNKNOWN);

      // Operands
      if (t == Token.REF)
      {
        CellReference cr = new CellReference(relativeTo);
        pos += cr.read(tokenData, pos);
        tokenStack.push(cr);
      }
      else if (t == Token.REFERR)
      {
        CellReferenceError cr = new CellReferenceError();
        pos += cr.read(tokenData, pos);
        tokenStack.push(cr);
      }
      else if (t == Token.ERR)
      {
        ErrorConstant ec = new ErrorConstant();
        pos += ec.read(tokenData, pos);
        tokenStack.push(ec);
      }
      else if (t == Token.REFV)
      {
        SharedFormulaCellReference cr = 
          new SharedFormulaCellReference(relativeTo);
        pos += cr.read(tokenData, pos);
        tokenStack.push(cr);
      }
      else if (t == Token.REF3D)
      {
        CellReference3d cr = new CellReference3d(relativeTo, workbook);
        pos += cr.read(tokenData, pos);
        tokenStack.push(cr);
      }
      else if (t == Token.AREA)
      {
        Area a = new Area();
        pos += a.read(tokenData, pos);
        tokenStack.push(a);
      }
      else if (t == Token.AREAV)
      {
        SharedFormulaArea a = new SharedFormulaArea(relativeTo);
        pos += a.read(tokenData, pos);
        tokenStack.push(a);
      }
      else if (t == Token.AREA3D)
      {
        Area3d a = new Area3d(workbook);
        pos += a.read(tokenData, pos);
        tokenStack.push(a);
      }
      else if (t == Token.NAME)
      {
        Name n = new Name();
        pos += n.read(tokenData, pos);
        n.setParseContext(parseContext);
        tokenStack.push(n);
      }
      else if (t == Token.NAMED_RANGE)
      {
        NameRange nr = new NameRange(nameTable);
        pos += nr.read(tokenData, pos);
        nr.setParseContext(parseContext);
        tokenStack.push(nr);
      }
      else if (t == Token.INTEGER)
      {
        IntegerValue i = new IntegerValue();
        pos += i.read(tokenData, pos);
        tokenStack.push(i);
      }
      else if (t == Token.DOUBLE)
      {
        DoubleValue d = new DoubleValue();
        pos += d.read(tokenData, pos);
        tokenStack.push(d);
      }
      else if (t == Token.BOOL)
      {
        BooleanValue bv = new BooleanValue();
        pos += bv.read(tokenData, pos);
        tokenStack.push(bv);
      }
      else if (t == Token.STRING)
      {
        StringValue sv = new StringValue(settings);
        pos += sv.read(tokenData, pos);
        tokenStack.push(sv);
      }
      else if (t == Token.MISSING_ARG)
      {
        MissingArg ma = new MissingArg();
        pos += ma.read(tokenData, pos);
        tokenStack.push(ma);
      }

      // Unary Operators
      else if (t == Token.UNARY_PLUS)
      {
        UnaryPlus up = new UnaryPlus();
        pos += up.read(tokenData, pos);
        addOperator(up);
      }
      else if (t == Token.UNARY_MINUS)
      {
        UnaryMinus um = new UnaryMinus();
        pos += um.read(tokenData, pos);
        addOperator(um);
      }
      else if (t == Token.PERCENT)
      {
        Percent p = new Percent();
        pos += p.read(tokenData, pos);
        addOperator(p);
      }

      // Binary Operators
      else if (t == Token.SUBTRACT)
      {
        Subtract s = new Subtract();
        pos += s.read(tokenData, pos);
        addOperator(s);
      }
      else if (t == Token.ADD)
      {
        Add s = new Add();
        pos += s.read(tokenData, pos);
        addOperator(s);
      }
      else if (t == Token.MULTIPLY)
      {
        Multiply s = new Multiply();
        pos += s.read(tokenData, pos);
        addOperator(s);
      }
      else if (t == Token.DIVIDE)
      {
        Divide s = new Divide();
        pos += s.read(tokenData, pos);
        addOperator(s);
      }
      else if (t == Token.CONCAT)
      {
        Concatenate c = new Concatenate();
        pos += c.read(tokenData, pos);
        addOperator(c);
      }
      else if (t == Token.POWER)
      {
        Power p = new Power();
        pos += p.read(tokenData, pos);
        addOperator(p);
      }
      else if (t == Token.LESS_THAN)
      {
        LessThan lt = new LessThan();
        pos += lt.read(tokenData, pos);
        addOperator(lt);
      }
      else if (t == Token.LESS_EQUAL)
      {
        LessEqual lte = new LessEqual();
        pos += lte.read(tokenData, pos);
        addOperator(lte);
      }
      else if (t == Token.GREATER_THAN)
      {
        GreaterThan gt = new GreaterThan();
        pos += gt.read(tokenData, pos);
        addOperator(gt);
      }
      else if (t == Token.GREATER_EQUAL)
      {
        GreaterEqual gte = new GreaterEqual();
        pos += gte.read(tokenData, pos);
        addOperator(gte);
      }
      else if (t == Token.NOT_EQUAL)
      {
        NotEqual ne = new NotEqual();
        pos += ne.read(tokenData, pos);
        addOperator(ne);
      }
      else if (t == Token.EQUAL)
      {
        Equal e = new Equal();
        pos += e.read(tokenData, pos);
        addOperator(e);
      }
      else if (t == Token.PARENTHESIS)
      {
        Parenthesis p = new Parenthesis();
        pos += p.read(tokenData, pos);
        addOperator(p);
      }

      // Functions
      else if (t == Token.ATTRIBUTE)
      {
        Attribute a = new Attribute(settings);
        pos += a.read(tokenData, pos);

        if (a.isSum())
        {
          addOperator(a);
        }
        else if (a.isIf())
        {
          // Add it to a special stack for ifs
          ifStack.push(a);
        }
      }
      else if (t == Token.FUNCTION)
      {
        BuiltInFunction bif = new BuiltInFunction(settings);
        pos += bif.read(tokenData, pos);

        addOperator(bif);
      }
      else if (t == Token.FUNCTIONVARARG)
      {
        VariableArgFunction vaf = new VariableArgFunction(settings);
        pos += vaf.read(tokenData, pos);

        if (vaf.getFunction() != Function.ATTRIBUTE)
        {
          addOperator(vaf);
        }
        else
        {
          // This is part of an IF function.  Get the operands, but then
          // add it to the top of the if stack
          vaf.getOperands(tokenStack);

          Attribute ifattr = null;
          if (ifStack.empty())
          {
            ifattr = new Attribute(settings);
          }
          else
          {
            ifattr = (Attribute) ifStack.pop();
          }
          
          ifattr.setIfConditions(vaf);
          tokenStack.push(ifattr);
        }
      }

      // Other things
      else if (t == Token.MEM_FUNC)
      {
        MemFunc memFunc = new MemFunc();
        handleMemoryFunction(memFunc);
      }
      else if (t == Token.MEM_AREA)
      {
        MemArea memArea = new MemArea();
        handleMemoryFunction(memArea);
      }
    }
  }

  /**
   * Handles a memory function
   */
  private void handleMemoryFunction(SubExpression subxp) 
    throws FormulaException
  {
    pos += subxp.read(tokenData, pos);

    // Create new tokenStack for the sub expression
    Stack oldStack = tokenStack;
    tokenStack = new Stack();

    parseSubExpression(subxp.getLength());

    ParseItem[] subexpr = new ParseItem[tokenStack.size()];
    int i = 0;
    while (!tokenStack.isEmpty())
    {
      subexpr[i] = (ParseItem) tokenStack.pop();
      i++;
    }

    subxp.setSubExpression(subexpr);
        
    tokenStack = oldStack;
    tokenStack.push(subxp);
  }

  /**
   * Adds the specified operator to the parse tree, taking operands off
   * the stack as appropriate
   */
  private void addOperator(Operator o)
  {
    // Get the operands off the stack
    o.getOperands(tokenStack);

    // Add this operator onto the stack
    tokenStack.push(o);
  }

  /**
   * Gets the formula as a string
   */
  public String getFormula()
  {
    StringBuffer sb = new StringBuffer();
    root.getString(sb);
    return sb.toString();
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
    root.adjustRelativeCellReferences(colAdjust, rowAdjust);
  }

  /**
   * Gets the bytes for the formula. This takes into account any
   * token mapping necessary because of shared formulas
   *
   * @return the bytes in RPN
   */
  public byte[] getBytes()
  {
    return root.getBytes();
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
    root.columnInserted(sheetIndex, col, currentSheet);
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
  public void columnRemoved(int sheetIndex, int col, boolean currentSheet)
  {
    root.columnRemoved(sheetIndex, col, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was inserted
   * @param row the column number which was inserted
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowInserted(int sheetIndex, int row, boolean currentSheet)
  {
    root.rowInserted(sheetIndex, row, currentSheet);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Tells
   * the formula  parser to update all of its cell references beyond this
   * column
   *
   * @param sheetIndex the sheet on which the column was removed
   * @param row the column number which was removed
   * @param currentSheet TRUE if this formula is on the sheet in which the
   * column was inserted, FALSE otherwise
   */
  public void rowRemoved(int sheetIndex, int row, boolean currentSheet)
  {
    root.rowRemoved(sheetIndex, row, currentSheet);
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   *
   * @return TRUE if the formula is valid import, FALSE otherwise
   */
  public boolean handleImportedCellReferences()
  {
    root.handleImportedCellReferences();
    return root.isValid();
  }
}
