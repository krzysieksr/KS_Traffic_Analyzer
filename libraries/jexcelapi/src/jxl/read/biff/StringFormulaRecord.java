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

import jxl.common.Assert;
import jxl.common.Logger;


import jxl.CellType;
import jxl.LabelCell;
import jxl.StringFormulaCell;
import jxl.WorkbookSettings;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WorkbookMethods;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;

/**
 * A string formula's last calculated value
 */
class StringFormulaRecord extends CellValue
  implements LabelCell, FormulaData, StringFormulaCell
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(StringFormulaRecord.class);

  /**
   * The last calculated value of the formula
   */
  private String value;

  /**
   * A handle to the class needed to access external sheets
   */
  private ExternalSheet externalSheet;

  /**
   * A handle to the name table
   */
  private WorkbookMethods nameTable;

  /**
   * The formula as an excel string
   */
  private String formulaString;

  /**
   * The raw data
   */
  private byte[] data;

  /**
   * Constructs this object from the raw data.  We need to use the excelFile
   * to retrieve the String record which follows this formula record
   *
   * @param t the raw data
   * @param excelFile the excel file
   * @param fr the formatting records
   * @param es the external sheet records
   * @param nt the workbook
   * @param si the sheet impl
   * @param ws the workbook settings
   */
  public StringFormulaRecord(Record t, File excelFile,
                             FormattingRecords fr,
                             ExternalSheet es,
                             WorkbookMethods nt,
                             SheetImpl si,
                             WorkbookSettings ws)
  {
    super(t, fr, si);

    externalSheet = es;
    nameTable = nt;

    data = getRecord().getData();

    int pos = excelFile.getPos();

    // Look for the string record in one of the records after the
    // formula.  Put a cap on it to prevent looping

    Record nextRecord = excelFile.next();
    int count = 0;
    while (nextRecord.getType() != Type.STRING && count < 4)
		{
      nextRecord = excelFile.next();
      count++;
    }
    Assert.verify(count < 4, " @ " + pos);
		byte[] stringData = nextRecord.getData();

		// Read in any continuation records
		nextRecord = excelFile.peek();
		while (nextRecord.getType() == Type.CONTINUE)
		{
			nextRecord = excelFile.next(); // move the pointer within the data
			byte[] d = new byte[stringData.length + nextRecord.getLength() - 1];
			System.arraycopy(stringData, 0, d, 0, stringData.length);
			System.arraycopy(nextRecord.getData(), 1, d, 
											 stringData.length, nextRecord.getLength() - 1);
			stringData = d;
			nextRecord = excelFile.peek();
		}
    readString(stringData, ws);
  }

  /**
   * Constructs this object from the raw data.  Used when reading in formula
   * strings which evaluate to null (in the case of some IF statements)
   *
   * @param t the raw data
   * @param fr the formatting records
   * @param es the external sheet records
   * @param nt the workbook
   * @param si the sheet impl
   * @param ws the workbook settings
   */
  public StringFormulaRecord(Record t,
                             FormattingRecords fr,
                             ExternalSheet es,
                             WorkbookMethods nt,
                             SheetImpl si)
  {
    super(t, fr, si);

    externalSheet = es;
    nameTable = nt;

    data = getRecord().getData();
    value = "";
  }


  /**
   * Reads in the string
   *
   * @param d the data
   * @param ws the workbook settings
   */
  private void readString(byte[] d, WorkbookSettings ws)
  {
    int pos = 0;
    int chars = IntegerHelper.getInt(d[0], d[1]);

    if (chars == 0)
    {
      value="";
      return;
    }
    pos += 2;
    int optionFlags = d[pos];
    pos++;

    if ((optionFlags & 0xf) != optionFlags)
    {
      // Uh oh - looks like a plain old string, not unicode
      // Recalculate all the positions
      pos = 0;
      chars = IntegerHelper.getInt(d[0], (byte) 0);
      optionFlags = d[1];
      pos = 2;
    }

    // See if it is an extended string
    boolean extendedString = ((optionFlags & 0x04) != 0);

    // See if string contains formatting information
    boolean richString = ((optionFlags & 0x08) != 0);

    if (richString)
    {
      pos += 2;
    }

    if (extendedString)
    {
      pos += 4;
    }

    // See if string is ASCII (compressed) or unicode
    boolean asciiEncoding = ((optionFlags & 0x01) == 0);

    if (asciiEncoding)
    {
      value = StringHelper.getString(d, chars, pos, ws);
    }
    else
    {
      value = StringHelper.getUnicodeString(d, chars, pos);
    }
  }

  /**
   * Interface method which returns the value
   *
   * @return the last calculated value of the formula
   */
  public String getContents()
  {
    return value;
  }

  /**
   * Interface method which returns the value
   *
   * @return the last calculated value of the formula
   */
  public String getString()
  {
    return value;
  }

  /**
   * Returns the cell type
   *
   * @return The cell type
   */
  public CellType getType()
  {
    return CellType.STRING_FORMULA;
  }

  /**
   * Gets the raw bytes for the formula.  This will include the
   * parsed tokens array
   *
   * @return the raw record data
   */
  public byte[] getFormulaData() throws FormulaException
  {
    if (!getSheet().getWorkbook().getWorkbookBof().isBiff8())
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

}
