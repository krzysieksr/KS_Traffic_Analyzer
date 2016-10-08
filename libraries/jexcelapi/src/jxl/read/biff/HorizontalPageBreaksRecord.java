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

import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * Contains the cell dimensions of this worksheet
 */
class HorizontalPageBreaksRecord extends RecordData
{
  /**
   * The logger
   */
  private final Logger logger = Logger.getLogger
    (HorizontalPageBreaksRecord.class);

  /**
   * The row page breaks
   */
  private int[] rowBreaks;

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static Biff7 biff7 = new Biff7();

  /**
   * Constructs the dimensions from the raw data
   *
   * @param t the raw data
   */
  public HorizontalPageBreaksRecord(Record t)
  {
    super(t);

    byte[] data = t.getData();

    int numbreaks = IntegerHelper.getInt(data[0], data[1]);
    int pos = 2;
    rowBreaks = new int[numbreaks];

    for (int i = 0; i < numbreaks; i++)
    {
      rowBreaks[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
      pos += 6;
    }
  }

  /**
   * Constructs the dimensions from the raw data
   *
   * @param t the raw data
   * @param biff7 an indicator to initialise this record for biff 7 format
   */
  public HorizontalPageBreaksRecord(Record t, Biff7 biff7)
  {
    super(t);

    byte[] data = t.getData();
    int numbreaks = IntegerHelper.getInt(data[0], data[1]);
    int pos = 2;
    rowBreaks = new int[numbreaks];
    for (int i = 0; i < numbreaks; i++)
    {
      rowBreaks[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
      pos += 2;
    }
  }

  /**
   * Gets the row breaks
   *
   * @return the row breaks on the current sheet
   */
  public int[] getRowBreaks()
  {
    return rowBreaks;
  }
}







