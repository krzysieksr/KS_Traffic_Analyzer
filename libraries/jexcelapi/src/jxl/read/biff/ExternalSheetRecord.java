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

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * An Externsheet record, containing the details of externally references
 * workbooks
 */
public class ExternalSheetRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(ExternalSheetRecord.class);

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static Biff7 biff7 = new Biff7();

  /**
   * An XTI structure
   */
  private static class XTI
  {
    /**
     * the supbook index
     */
    int supbookIndex;
    /**
     * the first tab
     */
    int firstTab;
    /**
     * the last tab
     */
    int lastTab;

    /**
     * Constructor
     *
     * @param s the supbook index
     * @param f the first tab
     * @param l the last tab
     */
    XTI(int s, int f, int l)
    {
      supbookIndex = s;
      firstTab = f;
      lastTab = l;
    }
  }

  /**
   * The array of XTI structures
   */
  private XTI[] xtiArray;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param ws the workbook settings
   */
  ExternalSheetRecord(Record t, WorkbookSettings ws)
  {
    super(t);
    byte[] data = getRecord().getData();

    int numxtis = IntegerHelper.getInt(data[0], data[1]);

    if (data.length < numxtis * 6 + 2)
    {
      xtiArray = new XTI[0];
      logger.warn("Could not process external sheets.  Formulas may " +
                  "be compromised.");
      return;
    }

    xtiArray = new XTI[numxtis];

    int pos = 2;
    for (int i = 0; i < numxtis; i++)
    {
      int s = IntegerHelper.getInt(data[pos],   data[pos + 1]);
      int f = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
      int l = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
      xtiArray[i] = new XTI(s, f, l);
      pos += 6;
    }
  }

  /**
   * Constructs this object from the raw data in biff 7 format.
   * Does nothing here
   *
   * @param t the raw data
   * @param settings the workbook settings
   * @param dummy dummy override to identify biff7 funcionality
   */
  ExternalSheetRecord(Record t, WorkbookSettings settings, Biff7 dummy)
  {
    super(t);

    logger.warn("External sheet record for Biff 7 not supported");
  }

  /**
   * Accessor for  the number of external sheet records
   * @return the number of XTI records
   */
  public int getNumRecords()
  {
    return xtiArray != null ? xtiArray.length : 0;
  }
  /**
   * Gets the supbook index for the specified external sheet
   *
   * @param index the index of the supbook record
   * @return the supbook index
   */
  public int getSupbookIndex(int index)
  {
    return xtiArray[index].supbookIndex;
  }

  /**
   * Gets the first tab index for the specified external sheet
   *
   * @param index the index of the supbook record
   * @return the first tab index
   */
  public int getFirstTabIndex(int index)
  {
    return xtiArray[index].firstTab;
  }

  /**
   * Gets the last tab index for the specified external sheet
   *
   * @param index the index of the supbook record
   * @return the last tab index
   */
  public int getLastTabIndex(int index)
  {
    return xtiArray[index].lastTab;
  }

  /**
   * Used when copying a workbook to access the raw external sheet data
   *
   * @return the raw external sheet data
   */
  public byte[] getData()
  {
    return getRecord().getData();
  }
}









