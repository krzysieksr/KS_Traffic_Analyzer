/*************************************************************************
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

package jxl.write.biff;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.CellReferenceHelper;
import jxl.CellType;
import jxl.FormulaCell;
import jxl.Sheet;
import jxl.WorkbookSettings;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WorkbookMethods;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;
import jxl.write.WritableCell;

/**
 * A formula record.  This is invoked when copying a formula from a
 * read only spreadsheet
 * This method implements the FormulaData interface to allow the copying
 * of writable sheets
 */
class ReadFormulaRecord extends CellValue implements FormulaData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(ReadFormulaRecord.class);

  /**
   * The underlying formula from the read sheet
   */
  private FormulaData formula;

  /**
   * The formula parser
   */
  private FormulaParser parser;

  /**
   * Constructor
   * 
   * @param f the formula to copy
   */
  protected ReadFormulaRecord(FormulaData f)
  {
    super(Type.FORMULA, f);
    formula = f;
  }

  protected final byte[] getCellData()
  {
    return super.getData();
  }

  /**
   * An exception has occurred, so produce some appropriate dummy
   * cell contents.  This may be overridden by subclasses
   * if they require specific handling
   *
   * @return the bodged data
   */
  protected byte[] handleFormulaException()
  {
    byte[] expressiondata = null;
    byte[] celldata = super.getData();

    // Generate an appropriate dummy formula
    WritableWorkbookImpl w = getSheet().getWorkbook();
    parser = new FormulaParser(getContents(), w, w, 
                               w.getSettings());

    // Get the bytes for the dummy formula
    try
    {
      parser.parse();
    }
    catch(FormulaException e2)
    {
      logger.warn(e2.getMessage());
      parser = new FormulaParser("\"ERROR\"", w, w, w.getSettings());
      try {parser.parse();} 
      catch(FormulaException e3) {Assert.verify(false);}
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
    return data;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    // Take the superclass cell data to take into account cell 
    // rationalization
    byte[] celldata = super.getData();
    byte[] expressiondata = null;

    try
    {
      if (parser == null)
      {
        expressiondata = formula.getFormulaData();
      }
      else
      {
        byte[]  formulaBytes = parser.getBytes();
        expressiondata = new byte[formulaBytes.length + 16];
        IntegerHelper.getTwoBytes(formulaBytes.length, expressiondata, 14);
        System.arraycopy(formulaBytes, 0, expressiondata, 16, 
                         formulaBytes.length);
      }

      // Set the recalculate on load bit
      expressiondata[8] |= 0x02;
    
      byte[] data = new byte[celldata.length + 
                             expressiondata.length];
      System.arraycopy(celldata, 0, data, 0, celldata.length);
      System.arraycopy(expressiondata, 0, data, 
                       celldata.length, expressiondata.length);
      return data;
    }
    catch (FormulaException e)
    {
      // Something has gone wrong trying to read the formula data eg. it
      // might be unsupported biff7 data
      logger.warn
        (CellReferenceHelper.getCellReference(getColumn(), getRow()) + 
         " " + e.getMessage());
      return handleFormulaException();
    }
  }

  /**
   * Returns the content type of this cell
   * 
   * @return the content type for this cell
   */
  public CellType getType()
  {
    return formula.getType();
  }

  /**
   * Quick and dirty function to return the contents of this cell as a string.
   *
   * @return the contents of this cell as a string
   */
  public String getContents()
  {
    return formula.getContents();
  }

  /**
   * Gets the raw bytes for the formula.  This will include the
   * parsed tokens array.  Used when copying spreadsheets
   *
   * @return the raw record data
   */
  public byte[] getFormulaData() throws FormulaException
  {
    byte[] d = formula.getFormulaData();
    byte[] data = new byte[d.length];

    System.arraycopy(d, 0, data, 0, d.length);
      
    // Set the recalculate on load bit
    data[8] |= 0x02;
    
    return data;
  }

  /**
   * Gets the formula bytes
   *
   * @return the formula bytes
   */
  public byte[] getFormulaBytes() throws FormulaException
  {
    // If the formula has been parsed, then get the parsed bytes
    if (parser != null)
    {
      return parser.getBytes();
    }

    // otherwise get the bytes from the original formula
    byte[] readFormulaData = getFormulaData();
    byte[] formulaBytes = new byte[readFormulaData.length - 16];
    System.arraycopy(readFormulaData, 16, formulaBytes, 0, 
                     formulaBytes.length);
    return formulaBytes;
  }

  /**
   * Implementation of the deep copy function
   *
   * @param col the column which the new cell will occupy
   * @param row the row which the new cell will occupy
   * @return  a copy of this cell, which can then be added to the sheet
   */
  public WritableCell copyTo(int col, int row)
  {
    return new FormulaRecord(col, row, this);
  }

  /**
   * Overrides the method in the base class to add this to the Workbook's
   * list of maintained formulas
   * 
   * @param fr the formatting records
   * @param ss the shared strings used within the workbook
   * @param s the sheet this is being added to
   */
  void setCellDetails(FormattingRecords fr, SharedStrings ss, 
                      WritableSheetImpl s)
  {
    super.setCellDetails(fr, ss, s);
    s.getWorkbook().addRCIRCell(this);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param col the column number which was inserted
   */
  void columnInserted(Sheet s, int sheetIndex, int col)
  {
    try
    {
      if (parser == null)
      {
        byte[] formulaData = formula.getFormulaData();
        byte[] formulaBytes = new byte[formulaData.length - 16];
        System.arraycopy(formulaData, 16, 
                         formulaBytes, 0, formulaBytes.length);
        parser = new FormulaParser(formulaBytes, 
                                   this, 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbookSettings());
        parser.parse();
      }

      parser.columnInserted(sheetIndex, col, s == getSheet());
    }
    catch (FormulaException e)
    {
      logger.warn("cannot insert column within formula:  " + e.getMessage());
    }
  }

  /**
   * Called when a column is removed on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param col the column number which was inserted
   */
  void columnRemoved(Sheet s, int sheetIndex, int col)
  {
    try
    {
      if (parser == null)
      {
        byte[] formulaData = formula.getFormulaData();
        byte[] formulaBytes = new byte[formulaData.length - 16];
        System.arraycopy(formulaData, 16, 
                         formulaBytes, 0, formulaBytes.length);
        parser = new FormulaParser(formulaBytes, 
                                   this, 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbookSettings());
        parser.parse();
      }

      parser.columnRemoved(sheetIndex, col, s == getSheet());
    }
    catch (FormulaException e)
    {
      logger.warn("cannot remove column within formula:  " + e.getMessage());
    }
  }

  /**
   * Called when a row is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the column was inserted
   * @param sheetIndex the sheet index on which the column was inserted
   * @param row the column number which was inserted
   */
  void rowInserted(Sheet s, int sheetIndex, int row)
  {
    try
    {
      if (parser == null)
      {
        byte[] formulaData = formula.getFormulaData();
        byte[] formulaBytes = new byte[formulaData.length - 16];
        System.arraycopy(formulaData, 16, 
                         formulaBytes, 0, formulaBytes.length);
        parser = new FormulaParser(formulaBytes, 
                                   this, 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbookSettings());
        parser.parse();
      }

      parser.rowInserted(sheetIndex, row, s == getSheet());
    }
    catch (FormulaException e)
    {
      logger.warn("cannot insert row within formula:  " + e.getMessage());
    }
  }

  /**
   * Called when a row is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change. The default implementation here does nothing
   *
   * @param s the sheet on which the row was removed
   * @param sheetIndex the sheet index on which the column was removed
   * @param row the column number which was removed
   */
  void rowRemoved(Sheet s, int sheetIndex, int row)
  {
    try
    {
      if (parser == null)
      {
        byte[] formulaData = formula.getFormulaData();
        byte[] formulaBytes = new byte[formulaData.length - 16];
        System.arraycopy(formulaData, 16, 
                         formulaBytes, 0, formulaBytes.length);
        parser = new FormulaParser(formulaBytes, 
                                   this, 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbook(), 
                                   getSheet().getWorkbookSettings());
        parser.parse();
      }

      parser.rowRemoved(sheetIndex, row, s == getSheet());
    }
    catch (FormulaException e)
    {
      logger.warn("cannot remove row within formula:  " + e.getMessage());
    }
  }

  /**
   * Accessor for the read formula
   *
   * @return the read formula
   */
  protected FormulaData getReadFormula()
  {
    return formula;
  }

  /**
   * Accessor for the read formula
   *
   * @return the read formula
   */
  public String getFormula() throws FormulaException
  {
    return ( (FormulaCell) formula).getFormula();
  }

  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * 
   * @return TRUE if this formula was able to be imported, FALSE otherwise
   */
  public boolean handleImportedCellReferences(ExternalSheet es,
                                              WorkbookMethods mt,
                                              WorkbookSettings ws)
  {
    try
    {
      if (parser == null)
      {
        byte[] formulaData = formula.getFormulaData();
        byte[] formulaBytes = new byte[formulaData.length - 16];
        System.arraycopy(formulaData, 16, 
                         formulaBytes, 0, formulaBytes.length);
        parser = new FormulaParser(formulaBytes, 
                                   this, 
                                   es, mt, ws);
        parser.parse();
      }

      return parser.handleImportedCellReferences();
    }
    catch (FormulaException e)
    {
      logger.warn("cannot import formula:  " + e.getMessage());
      return false;
    }
  }
}
