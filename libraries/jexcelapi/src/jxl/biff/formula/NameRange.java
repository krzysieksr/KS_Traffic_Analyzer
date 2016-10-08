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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.biff.IntegerHelper;
import jxl.biff.NameRangeException;
import jxl.biff.WorkbookMethods;

/**
 * A name operand
 */
class NameRange extends Operand implements ParsedThing
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(NameRange.class);

  /**
   * A handle to the name table
   */
  private WorkbookMethods nameTable;

  /**
   * The string name
   */
  private String name;
  
  /**
   * The index into the name table
   */
  private int index;

  /**
   * Constructor
   */
  public NameRange(WorkbookMethods nt)
  {
    nameTable = nt;
    Assert.verify(nameTable != null);
  }

  /**
   * Constructor when parsing a string via the api
   * 
   * @param nm the name string
   * @param nt the name table
   */
  public NameRange(String nm, WorkbookMethods nt) throws FormulaException
  {
    name = nm;
    nameTable = nt;

    index = nameTable.getNameIndex(name);

    if (index < 0 )
    {
      throw new FormulaException(FormulaException.CELL_NAME_NOT_FOUND, name);
    }

    index += 1; // indexes are 1-based
  }

  /** 
   * Reads the ptg data from the array starting at the specified position
   *
   * @param data the RPN array
   * @param pos the current position in the array, excluding the ptg identifier
   * @return the number of bytes read
   */
  public int read(byte[] data, int pos) throws FormulaException 
  {
    try
    {
      index = IntegerHelper.getInt(data[pos], data[pos+1]);

      name = nameTable.getName(index - 1); // ilbl is 1-based
      
      return 4;
    }
    catch (NameRangeException e)
    {
      throw new FormulaException(FormulaException.CELL_NAME_NOT_FOUND,"");
    }
  }

  /**
   * Gets the token representation of this item in RPN
   *
   * @return the bytes applicable to this formula
   */
  byte[] getBytes()
  {
    byte[] data = new byte[5];

    data[0] = Token.NAMED_RANGE.getValueCode();
    
    if (getParseContext() == ParseContext.DATA_VALIDATION)
    {
      data[0] = Token.NAMED_RANGE.getReferenceCode();
    }

    IntegerHelper.getTwoBytes(index, data, 1);

    return data;
  }

  /**
   * Abstract method implementation to get the string equivalent of this
   * token
   * 
   * @param buf the string to append to
   */
  public void getString(StringBuffer buf)
  {
    buf.append(name);
  }


  /**
   * If this formula was on an imported sheet, check that
   * cell references to another sheet are warned appropriately
   * Flags the formula as invalid
   */
  void handleImportedCellReferences()
  {
    setInvalid();
  }

}
