/*********************************************************************
*
*      Copyright (C) 2005 Andrew Khan
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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.BooleanCell;
import jxl.BooleanFormulaCell;
import jxl.CellType;
import jxl.biff.DoubleHelper;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.IntegerHelper;
import jxl.biff.WorkbookMethods;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;

/**
 * A shared boolean formula record
 */
public class SharedBooleanFormulaRecord extends BaseSharedFormulaRecord
  implements BooleanCell, FormulaData, BooleanFormulaCell
{
  /**
   * The logger
   */
  private static Logger logger = 
    Logger.getLogger(SharedBooleanFormulaRecord.class);

  /**
   * The boolean value of this cell.  If this cell represents an error,
   * this will be false
   */
  private boolean value;

  /**
   * Constructs this number
   *
   * @param t the data
   * @param excelFile the excel biff data
   * @param v the value
   * @param fr the formatting records
   * @param es the external sheet
   * @param nt the name table
   * @param si the sheet
   */
  public SharedBooleanFormulaRecord(Record t,
                                    File excelFile,
                                    boolean v,
                                    FormattingRecords fr,
                                    ExternalSheet es,
                                    WorkbookMethods nt,
                                    SheetImpl si)
  {
    super(t, fr, es, nt, si, excelFile.getPos());
    value = v;
  }

  /**
   * Interface method which Gets the boolean value stored in this cell.  If
   * this cell contains an error, then returns FALSE.  Always query this cell
   *  type using the accessor method isError() prior to calling this method
   *
   * @return TRUE if this cell contains TRUE, FALSE if it contains FALSE or
   * an error code
   */
  public boolean getValue()
  {
    return value;
  }

  /**
   * Returns the numerical value as a string
   *
   * @return The numerical value of the formula as a string
   */
  public String getContents()
  {
    // return Boolean.toString(value) - only available in 1.4 or later
    return (new Boolean(value)).toString();
  }

  /**
   * Returns the cell type
   *
   * @return The cell type
   */
  public CellType getType()
  {
    return CellType.BOOLEAN_FORMULA;
  }

  /**
   * Gets the raw bytes for the formula.  This will include the
   * parsed tokens array.  Used when copying spreadsheets
   *
   * @return the raw record data
   * @exception FormulaException
   */
  public byte[] getFormulaData() throws FormulaException
  {
    if (!getSheet().getWorkbookBof().isBiff8())
    {
      throw new FormulaException(FormulaException.BIFF8_SUPPORTED);
    }

    // Get the tokens, taking into account the mapping from shared
    // formula specific values into normal values
    FormulaParser fp = new FormulaParser
      (getTokens(), this,
       getExternalSheet(), getNameTable(),
       getSheet().getWorkbook().getSettings());
    fp.parse();
    byte[] rpnTokens = fp.getBytes();

    byte[] data = new byte[rpnTokens.length + 22];

    // Set the standard info for this cell
    IntegerHelper.getTwoBytes(getRow(), data, 0);
    IntegerHelper.getTwoBytes(getColumn(), data, 2);
    IntegerHelper.getTwoBytes(getXFIndex(), data, 4);
    data[6] = (byte) 1;
    data[8] = (byte) (value == true ? 1 : 0);
    data[12] = (byte) 0xff;
    data[13] = (byte) 0xff;

    // Now copy in the parsed tokens
    System.arraycopy(rpnTokens, 0, data, 22, rpnTokens.length);
    IntegerHelper.getTwoBytes(rpnTokens.length, data, 20);

    // Lop off the standard information
    byte[] d = new byte[data.length - 6];
    System.arraycopy(data, 6, d, 0, data.length - 6);

    return d;
  }
}
