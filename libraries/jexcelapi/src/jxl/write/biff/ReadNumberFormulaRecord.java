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

package jxl.write.biff;

import java.text.NumberFormat;

import jxl.common.Logger;

import jxl.NumberFormulaCell;
import jxl.biff.DoubleHelper;
import jxl.biff.FormulaData;
import jxl.biff.IntegerHelper;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;

/**
 * Class for read number formula records
 */
class ReadNumberFormulaRecord extends ReadFormulaRecord 
  implements NumberFormulaCell
{
  // The logger
  private static Logger logger = Logger.getLogger(ReadNumberFormulaRecord.class);

  /**
   * Constructor
   *
   * @param f
   */
  public ReadNumberFormulaRecord(FormulaData f)
  {
    super(f);
  }

  /**
   * Gets the double contents for this cell.
   *
   * @return the cell contents
   */
  public double getValue()
  {
    return ( (NumberFormulaCell) getReadFormula()).getValue();
  }

  /**
   * Gets the NumberFormat used to format this cell.  This is the java
   * equivalent of the Excel format
   *
   * @return the NumberFormat used to format the cell
   */
  public NumberFormat getNumberFormat()
  {
    return ( (NumberFormulaCell) getReadFormula()).getNumberFormat();
  }

  /**
   * Error formula specific exception handling.  Can't really create
   * a formula (as it will look for a cell of that name, so just
   * create a STRING record containing the contents
   *
   * @return the bodged data
   */
  protected byte[] handleFormulaException()
  {
    byte[] expressiondata = null;
    byte[] celldata = super.getCellData();

    // Generate an appropriate dummy formula
    WritableWorkbookImpl w = getSheet().getWorkbook();
    FormulaParser parser = new FormulaParser(Double.toString(getValue()), w, w,
                                             w.getSettings());

    // Get the bytes for the dummy formula
    try
    {
      parser.parse();
    }
    catch(FormulaException e2)
    {
      logger.warn(e2.getMessage());
    }
    byte[] formulaBytes = parser.getBytes();
    expressiondata = new byte[formulaBytes.length + 16];
    IntegerHelper.getTwoBytes(formulaBytes.length, expressiondata, 14);
    System.arraycopy(formulaBytes, 0, expressiondata, 16, 
                     formulaBytes.length);

    // Set the recalculate on load bit
    expressiondata[8] |= 0x02;
    
    byte[] data = new byte[celldata.length + 
                           expressiondata.length];
    System.arraycopy(celldata, 0, data, 0, celldata.length);
    System.arraycopy(expressiondata, 0, data, 
                     celldata.length, expressiondata.length);

    // Store the value in the formula
    DoubleHelper.getIEEEBytes(getValue(), data, 6);

    return data;
  }

}
