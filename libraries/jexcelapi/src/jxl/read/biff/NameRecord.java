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

import java.util.ArrayList;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.BuiltInName;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.StringHelper;

/**
 * Holds an excel name record, and the details of the cells/ranges it refers
 * to
 */
public class NameRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(NameRecord.class);

  /**
   * The name
   */
  private String name;

  /**
   * The built in name
   */
  private BuiltInName builtInName;

  /**
   * The 0-based index in the name table
   */
  private int index;
  
  /** 
   * The 0-based index sheet reference for a record name
   *     0 is for a global reference
   */
  private int sheetRef = 0;

  /**
   * Indicates whether this is a biff8 name record.  Used during copying
   */
  private boolean isbiff8;

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static Biff7 biff7 = new Biff7();

  // Constants which refer to the name type
  private static final int commandMacro = 0xc;
  private static final int builtIn = 0x20;

  // Constants which refer to the parse tokens after the string
  private static final int cellReference = 0x3a;
  private static final int areaReference = 0x3b;
  private static final int subExpression = 0x29;
  private static final int union         = 0x10;

  /**
   * A nested class to hold range information
   */
  public class NameRange
  {
    /**
     * The first column
     */
    private int columnFirst;

    /**
     * The first row
     */
    private int rowFirst;

    /**
     * The last column
     */
    private int columnLast;

    /**
     * The last row
     */
    private int rowLast;

    /**
     * The first sheet
     */
    private int externalSheet;

    /**
     * Constructor
     *
     * @param s1 the sheet
     * @param c1 the first column
     * @param r1 the first row
     * @param c2 the last column
     * @param r2 the last row
     */
    NameRange(int s1, int c1, int r1, int c2, int r2)
    {
      columnFirst = c1;
      rowFirst = r1;
      columnLast = c2;
      rowLast = r2;
      externalSheet = s1;
    }

    /**
     * Accessor for the first column
     *
     * @return the index of the first column
     */
    public int getFirstColumn()
    {
      return columnFirst;
    }

    /**
     * Accessor for the first row
     *
     * @return the index of the first row
     */
    public int getFirstRow()
    {
      return rowFirst;
    }

    /**
     * Accessor for the last column
     *
     * @return the index of the last column
     */
    public int getLastColumn()
    {
      return columnLast;
    }

    /**
     * Accessor for the last row
     *
     * @return the index of the last row
     */
    public int getLastRow()
    {
      return rowLast;
    }

    /**
     * Accessor for the first sheet
     *
     * @return  the index of the external  sheet
     */
    public int getExternalSheet()
    {
      return externalSheet;
    }
  }

  /**
   * The ranges referenced by this name
   */
  private ArrayList ranges;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param ws the workbook settings
   * @param ind the index in the name table
   */
  NameRecord(Record t, WorkbookSettings ws, int ind)
  {
    super(t);
    index = ind;
    isbiff8 = true;

    try
    {
      ranges = new ArrayList();

      byte[] data = getRecord().getData();
      int option = IntegerHelper.getInt(data[0], data[1]);
      int length = data[3];
      sheetRef = IntegerHelper.getInt(data[8],data[9]);

      if ((option & builtIn) != 0)
      {
        builtInName = BuiltInName.getBuiltInName(data[15]);
      }
      else
      {
        name = StringHelper.getString(data, length, 15, ws);
      }

      if ((option & commandMacro) != 0)
      {
        // This is a command macro, so it has no cell references
        return;
      }

      int pos = length + 15;
      
      if (data[pos] == cellReference)
      {
        int sheet = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
        int row = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
        int columnMask = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);
        int column = columnMask & 0xff;

        // Check that we are not dealing with offsets
        Assert.verify((columnMask & 0xc0000) == 0);

        NameRange r = new NameRange(sheet, column, row, column, row);
        ranges.add(r);
      }
      else if (data[pos] == areaReference)
      {
        int sheet1 = 0;
        int r1 = 0;
        int columnMask = 0;
        int c1 = 0;
        int r2 = 0;
        int c2 = 0;
        NameRange range = null;

        while (pos < data.length)
        {
          sheet1 = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
          r1 = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
          r2 = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);

          columnMask = IntegerHelper.getInt(data[pos + 7], data[pos + 8]);
          c1 = columnMask & 0xff;

          // Check that we are not dealing with offsets
          Assert.verify((columnMask & 0xc0000) == 0);

          columnMask = IntegerHelper.getInt(data[pos + 9], data[pos + 10]);
          c2 = columnMask & 0xff;

          // Check that we are not dealing with offsets
          Assert.verify((columnMask & 0xc0000) == 0);

          range = new NameRange(sheet1, c1, r1,  c2, r2);
          ranges.add(range);

          pos += 11;
        }
      }
      else if (data[pos] == subExpression)
      {
        int sheet1 = 0;
        int r1 = 0;
        int columnMask = 0;
        int c1 = 0;
        int r2 = 0;
        int c2 = 0;
        NameRange range = null;

        // Consume unnecessary parsed tokens
        if (pos < data.length &&
            data[pos] != cellReference &&
            data[pos] != areaReference)
        {
          if (data[pos] == subExpression)
          {
            pos += 3;
          }
          else if (data[pos] == union)
          {
            pos += 1;
          }
        }

        while (pos < data.length)
        {
          sheet1 = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
          r1 = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
          r2 = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);

          columnMask = IntegerHelper.getInt(data[pos + 7], data[pos + 8]);
          c1 = columnMask & 0xff;

          // Check that we are not dealing with offsets
          Assert.verify((columnMask & 0xc0000) == 0);

          columnMask = IntegerHelper.getInt(data[pos + 9], data[pos + 10]);
          c2 = columnMask & 0xff;

          // Check that we are not dealing with offsets
          Assert.verify((columnMask & 0xc0000) == 0);

          range = new NameRange(sheet1, c1, r1, c2, r2);
          ranges.add(range);

          pos += 11;

          // Consume unnecessary parsed tokens
          if (pos < data.length &&
              data[pos] != cellReference &&
              data[pos] != areaReference)
          {
            if (data[pos] == subExpression)
            {
              pos += 3;
            }
            else if (data[pos] == union)
            {
              pos += 1;
            }
          }
        }
      }
      else
      {
          String n = name != null ? name : builtInName.getName();
          logger.warn("Cannot read name ranges for " + n + 
                      " - setting to empty");
          NameRange range = new NameRange(0,0,0,0,0);
          ranges.add(range);
      }
    }
    catch (Throwable t1)
    {
      // Generate a warning
      // Names are really a nice to have, and we don't want to halt the
      // reading process for functionality that probably won't be used
      logger.warn("Cannot read name");
      name = "ERROR";
    }
  }

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param ws the workbook settings
   * @param ind the index in the name table
   * @param dummy dummy parameter to indicate a biff7 workbook
   */
  NameRecord(Record t, WorkbookSettings ws, int ind, Biff7 dummy)
  {
    super(t);
    index = ind;
    isbiff8 = false;

    try
    {
      ranges = new ArrayList();
      byte[] data = getRecord().getData();
      int length = data[3];
      sheetRef = IntegerHelper.getInt(data[8], data[9]);
      name = StringHelper.getString(data, length, 14, ws);

      int pos = length + 14;

      if (pos >= data.length)
      {
        // There appears to be nothing after the name, so return
        return;
      }

      if (data[pos] == cellReference)
      {
        int sheet = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
        int row = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
        int column = data[pos + 17];

        NameRange r = new NameRange(sheet, column, row, column, row);
        ranges.add(r);
      }
      else if (data[pos] == areaReference)
      {
        int sheet1 = 0;
        int r1 = 0;
        int c1 = 0;
        int r2 = 0;
        int c2 = 0;
        NameRange range = null;

        while (pos < data.length)
        {
          sheet1 = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
          r1 = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
          r2 = IntegerHelper.getInt(data[pos + 17], data[pos + 18]);

          c1 = data[pos + 19];
          c2 = data[pos + 20];

          range = new NameRange(sheet1, c1, r1, c2, r2);
          ranges.add(range);

          pos += 21;
        }
      }
      else if (data[pos] == subExpression)
      {
        int sheet1 = 0;
        int sheet2 = 0;
        int r1 = 0;
        int c1 = 0;
        int r2 = 0;
        int c2 = 0;
        NameRange range = null;

        // Consume unnecessary parsed tokens
        if (pos < data.length &&
            data[pos] != cellReference &&
            data[pos] != areaReference)
        {
          if (data[pos] == subExpression)
          {
            pos += 3;
          }
          else if (data[pos] == union)
          {
            pos += 1;
          }
        }

        while (pos < data.length)
        {
          sheet1 = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
          r1 = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
          r2 = IntegerHelper.getInt(data[pos + 17], data[pos + 18]);

          c1 = data[pos + 19];
          c2 = data[pos + 20];

          range = new NameRange(sheet1, c1, r1, c2, r2);
          ranges.add(range);

          pos += 21;

          // Consume unnecessary parsed tokens
          if (pos < data.length &&
              data[pos] != cellReference &&
              data[pos] != areaReference)
          {
            if (data[pos] == subExpression)
            {
              pos += 3;
            }
            else if (data[pos] == union)
            {
              pos += 1;
            }
          }
        }
      }
    }
    catch (Throwable t1)
    {
      // Generate a warning
      // Names are really a nice to have, and we don't want to halt the
      // reading process for functionality that probably won't be used
      logger.warn("Cannot read name.");
      name = "ERROR";
    }
  }

  /**
   * Gets the name
   *
   * @return the strings
   */
  public String getName()
  {
    return name;
  }

  /**
   * Gets the built in name
   *
   * @return the built in name
   */
  public BuiltInName getBuiltInName()
  {
    return builtInName;
  }

  /**
   * Gets the array of ranges for this name.  This method is public as it is
   * used from the writable side when copying ranges
   *
   * @return the ranges
   */
  public NameRange[] getRanges()
  {
    NameRange[] nr = new NameRange[ranges.size()];
    return (NameRange[]) ranges.toArray(nr);
  }

  /**
   * Accessor for the index into the name table
   *
   * @return the 0-based index into the name table
   */
  int getIndex()
  {
    return index;
  }
  
  /**
   * The 0-based index sheet reference for a record name
   *     0 is for a global reference
   *
   * @return the sheet reference for name formula
   */
  public int getSheetRef()
  {
    return sheetRef;
  }
  
  /**
   * Set the index sheet reference for a record name
   *     0 is for a global reference
   */
  public void setSheetRef(int i)
  {
    sheetRef = i;
  }

  /**
   * Called when copying a sheet.  Just returns the raw data
   *
   * @return the raw data
   */
  public byte[] getData()
  {
    return getRecord().getData();
  }
  
  /**
   * Called when copying to determine whether this is a biff8 name
   *
   * @return TRUE if this is a biff8 name record, FALSE otherwise
   */
  public boolean isBiff8()
  {
    return isbiff8;
  }

  /**
   * Queries whether this is a global name or not
   *
   * @return TRUE if this is a global name, FALSE otherwise
   */
  public boolean isGlobal()
  {
    return sheetRef == 0;
  }
}

