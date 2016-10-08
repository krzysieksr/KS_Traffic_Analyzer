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

package jxl.write.biff;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.EncodedURLHelper;
import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Stores the supporting workbook information.  For files written by
 * JExcelApi this will only reference internal sheets
 */
class SupbookRecord extends WritableRecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(SupbookRecord.class);

  /**
   * The type of this supbook record
   */
  private SupbookType type;

  /**
   * The data to be written to the binary file
   */
  private byte[] data;

  /**
   * The number of sheets - internal & external supbooks only
   */
  private int numSheets;

  /**
   * The name of the external file
   */
  private String fileName;

  /**
   * The names of the external sheets
   */
  private String[] sheetNames;

  /**
   * The workbook settings
   */
  private WorkbookSettings workbookSettings;

  /**
   * The type of supbook this refers to
   */
  private static class SupbookType {};

  public final static SupbookType INTERNAL = new SupbookType();
  public final static SupbookType EXTERNAL = new SupbookType();
  public final static SupbookType ADDIN    = new SupbookType();
  public final static SupbookType LINK     = new SupbookType();
  public final static SupbookType UNKNOWN  = new SupbookType();
  
  /**
   * Constructor for add in function names
   */
  public SupbookRecord()
  {
    super(Type.SUPBOOK);
    type = ADDIN;
  }

  /**
   * Constructor for internal sheets
   */
  public SupbookRecord(int sheets, WorkbookSettings ws)
  {
    super(Type.SUPBOOK);

    numSheets = sheets;
    type = INTERNAL;
    workbookSettings = ws;
  }

  /**
   * Constructor for external sheets
   *
   * @param fn the filename of the external supbook
   * @param ws the workbook settings
   */
  public SupbookRecord(String fn, WorkbookSettings ws)
  {
    super(Type.SUPBOOK);

    fileName = fn;
    numSheets = 1;
    sheetNames = new String[0];
    workbookSettings = ws;

    type = EXTERNAL;
  }

  /**
   * Constructor used when copying from an external workbook
   */
  public SupbookRecord(jxl.read.biff.SupbookRecord sr, WorkbookSettings ws)
  {
    super(Type.SUPBOOK);

    workbookSettings = ws;
    if (sr.getType() == sr.INTERNAL)
    {
      type = INTERNAL;
      numSheets = sr.getNumberOfSheets();
    }
    else if (sr.getType() == sr.EXTERNAL)
    {
      type = EXTERNAL;
      numSheets = sr.getNumberOfSheets();
      fileName = sr.getFileName();
      sheetNames = new String[numSheets];

      for (int i = 0; i < numSheets; i++)
      {
        sheetNames[i] = sr.getSheetName(i);
      }
    }

    if (sr.getType() == sr.ADDIN)
    {
      logger.warn("Supbook type is addin");
    }
  }

  /**
   * Initializes an internal supbook record
   * 
   * @param sr the read supbook record to copy from
   */
  private void initInternal(jxl.read.biff.SupbookRecord sr)
  {
    numSheets = sr.getNumberOfSheets();
    initInternal();
  }

  /**
   * Initializes an internal supbook record
   */
  private void initInternal()
  {
    data = new byte[4];

    IntegerHelper.getTwoBytes(numSheets, data, 0);
    data[2] = 0x1;
    data[3] = 0x4;    
    type = INTERNAL;
  }

  /**
   * Adjust the number of internal sheets.  Called by WritableSheet when
   * a sheet is added or or removed to the workbook
   *
   * @param sheets the new number of sheets
   */
  void adjustInternal(int sheets)
  {
    Assert.verify(type == INTERNAL);
    numSheets = sheets;
    initInternal();
  }

  /**
   * Initializes an external supbook record
   */
  private void initExternal()
  {
    int totalSheetNameLength = 0;
    for (int i = 0; i < numSheets; i++)
    {
      totalSheetNameLength += sheetNames[i].length();
    }

    byte[] fileNameData = EncodedURLHelper.getEncodedURL(fileName, 
                                                         workbookSettings);
    int dataLength = 2 + // numsheets
      4 + fileNameData.length +
      numSheets * 3 + totalSheetNameLength * 2;
      
    data = new byte[dataLength];

    IntegerHelper.getTwoBytes(numSheets, data, 0);
    
    // Add in the file name.  Precede with a byte denoting that it is a 
    // file name
    int pos = 2;
    IntegerHelper.getTwoBytes(fileNameData.length+1, data, pos);
    data[pos+2] = 0; // ascii indicator
    data[pos+3] = 1; // file name indicator
    System.arraycopy(fileNameData, 0, data, pos+4, fileNameData.length);

    pos += 4 + fileNameData.length;

    // Get the sheet names
    for (int i = 0; i < sheetNames.length; i++)
    {
      IntegerHelper.getTwoBytes(sheetNames[i].length(), data, pos);
      data[pos+2] = 1; // unicode indicator
      StringHelper.getUnicodeBytes(sheetNames[i], data, pos+3);
      pos += 3 + sheetNames[i].length() * 2;
    }
  }

  /**
   * Initializes the supbook record for add in functions
   */
  private void initAddin()
  {
    data = new byte[] {0x1, 0x0, 0x1, 0x3a};
  }

  /**
   * The binary data to be written out
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    if (type == INTERNAL)
    {
      initInternal();
    }
    else if (type == EXTERNAL)
    {
      initExternal();
    }
    else if (type == ADDIN)
    {
      initAddin();
    }
    else
    {
      logger.warn("unsupported supbook type - defaulting to internal");
      initInternal();
    }

    return data;
  }

  /**
   * Gets the type of this supbook record
   * 
   * @return the type of this supbook
   */
  public SupbookType getType()
  {
    return type;
  }

  /**
   * Gets the number of sheets.  This will only be non-zero for internal
   * and external supbooks
   *
   * @return the number of sheets
   */
  public int getNumberOfSheets()
  {
    return numSheets;
  }

  /**
   * Accessor for the file name
   *
   * @return the file name
   */
  public String getFileName()
  {
    return fileName;
  }

  /**
   * Adds the worksheet name to this supbook
   *
   * @param name the worksheet name
   * @return the index of this sheet in the supbook record
   */
  public int getSheetIndex(String s)
  {
    boolean found = false;
    int sheetIndex = 0;
    for (int i = 0; i < sheetNames.length && !found; i++)
    {
      if (sheetNames[i].equals(s))
      {
        found = true;
        sheetIndex = 0;
      }
    }

    if (found)
    {
      return sheetIndex;
    }

    // Grow the array
    String[] names = new String[sheetNames.length + 1];
    System.arraycopy(sheetNames, 0, names, 0, sheetNames.length);
    names[sheetNames.length] = s;
    sheetNames = names;
    return sheetNames.length - 1;
  }

  /**
   * Accessor for the sheet name
   * 
   * @param s the sheet index
   */
  public String getSheetName(int s)
  {
    return sheetNames[s];
  }
}
