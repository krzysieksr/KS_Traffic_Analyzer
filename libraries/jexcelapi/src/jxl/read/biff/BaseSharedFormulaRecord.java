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

package jxl.read.biff;

import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.WorkbookMethods;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;

/**
 * A base class for shared formula records
 */
public abstract class BaseSharedFormulaRecord extends CellValue
  implements FormulaData
{
  /**
   * The formula as an excel string
   */
  private String formulaString;

  /**
   * The position of the next record in the file.  Used when looking for
   * for subsequent records eg. a string value
   */
  private int filePos;

  /**
   * The array of parsed tokens
   */
  private byte[] tokens;

  /**
   * The external sheet
   */
  private ExternalSheet externalSheet;

  /**
   * The name table
   */
  private WorkbookMethods nameTable;

  /**
   * Constructs this number
   *
   * @param t the record
   * @param fr the formatting records
   * @param es the external sheet
   * @param nt the name table
   * @param si the sheet
   * @param pos the position of the next record in the file
   */
  public BaseSharedFormulaRecord(Record t,
                                 FormattingRecords fr,
                                 ExternalSheet es,
                                 WorkbookMethods nt,
                                 SheetImpl si,
                                 int pos)
  {
    super(t, fr, si);
    externalSheet = es;
    nameTable = nt;
    filePos = pos;
  }

  /**
   * Gets the formula as an excel string
   *
   * @return the formula as an excel string
   * @exception FormulaException
   */
  public String getFormula() throws FormulaException
  {
    if (formulaString == null)
    {
      FormulaParser fp = new FormulaParser
        (tokens, this, externalSheet, nameTable,
         getSheet().getWorkbook().getSettings());
      fp.parse();
      formulaString = fp.getFormula();
    }

    return formulaString;
  }

  /**
   * Called by the shared formula record to set the tokens for
   * this formula
   *
   * @param t the tokens
   */
  void setTokens(byte[] t)
  {
    tokens = t;
  }

  /**
   * Accessor for the tokens which make up this formula
   *
   * @return the tokens
   */
  protected final byte[] getTokens()
  {
    return tokens;
  }

  /**
   * Access for the external sheet
   *
   * @return the external sheet
   */
  protected final ExternalSheet getExternalSheet()
  {
    return externalSheet;
  }

  /**
   * Access for the name table
   *
   * @return the name table
   */
  protected final WorkbookMethods getNameTable()
  {
    return nameTable;
  }

  /**
   * In case the shared formula is not added for any reason, we need
   * to expose the raw record data , in order to try again
   *
   * @return the record data from the base class
   */
  public Record getRecord()
  {
    return super.getRecord();
  }

  /**
   * Accessor for the position of the next record
   *
   * @return the position of the next record
   */
  final int getFilePos()
  {
    return filePos;
  }
}









