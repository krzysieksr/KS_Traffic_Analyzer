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

import java.text.DecimalFormat;
import java.text.NumberFormat;

import jxl.common.Logger;

import jxl.CellType;
import jxl.NumberCell;
import jxl.NumberFormulaCell;
import jxl.biff.DoubleHelper;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.WorkbookMethods;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;

/**
 * A formula's last calculated value
 */
class NumberFormulaRecord extends CellValue
  implements NumberCell, FormulaData, NumberFormulaCell
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(NumberFormulaRecord.class);

  /**
   * The last calculated value of the formula
   */
  private double value;

  /**
   * The number format
   */
  private NumberFormat format;

  /**
   * The string format for the double value
   */
  private static final DecimalFormat defaultFormat =
    new DecimalFormat("#.###");

  /**
   * The formula as an excel string
   */
  private String formulaString;

  /**
   * A handle to the class needed to access external sheets
   */
  private ExternalSheet externalSheet;

  /**
   * A handle to the name table
   */
  private WorkbookMethods nameTable;

  /**
   * The raw data
   */
  private byte[] data;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param fr the formatting record
   * @param es the external sheet
   * @param nt the name table
   * @param si the sheet
   */
  public NumberFormulaRecord(Record t, FormattingRecords fr,
                             ExternalSheet es, WorkbookMethods nt,
                             SheetImpl si)
  {
    super(t, fr, si);

    externalSheet = es;
    nameTable = nt;
    data = getRecord().getData();

    format = fr.getNumberFormat(getXFIndex());

    if (format == null)
    {
      format = defaultFormat;
    }

    value = DoubleHelper.getIEEEDouble(data, 6);
  }

  /**
   * Interface method which returns the value
   *
   * @return the last calculated value of the formula
   */
  public double getValue()
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
    return !Double.isNaN(value) ? format.format(value) : "";
  }

  /**
   * Returns the cell type
   *
   * @return The cell type
   */
  public CellType getType()
  {
    return CellType.NUMBER_FORMULA;
  }

  /**
   * Gets the raw bytes for the formula.  This will include the
   * parsed tokens array.  Used when copying spreadsheets
   *
   * @return the raw record data
   */
  public byte[] getFormulaData() throws FormulaException
  {
    if (!getSheet().getWorkbookBof().isBiff8())
    {
      throw new FormulaException(FormulaException.BIFF8_SUPPORTED);
    }

    // Lop off the standard information
    byte[] d = new byte[data.length - 6];
    System.arraycopy(data, 6, d, 0, data.length - 6);

    return d;
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
      byte[] tokens = new byte[data.length - 22];
      System.arraycopy(data, 22, tokens, 0, tokens.length);
      FormulaParser fp = new FormulaParser
        (tokens, this, externalSheet, nameTable,
         getSheet().getWorkbook().getSettings());
      fp.parse();
      formulaString = fp.getFormula();
    }

    return formulaString;
  }

  /**
   * Gets the NumberFormat used to format this cell.  This is the java
   * equivalent of the Excel format
   *
   * @return the NumberFormat used to format the cell
   */
  public NumberFormat getNumberFormat()
  {
    return format;
  }
}
