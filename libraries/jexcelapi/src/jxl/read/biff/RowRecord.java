/*********************************************************************
*
*      Copyright (C) 2001 Andrew Khan
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
 * A row  record
 */
public class RowRecord extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(RowRecord.class);

  /**
   * The number of this row
   */
  private int rowNumber;
  /**
   * The height of this row
   */
  private int rowHeight;
  /**
   * Flag to indicate whether this row is collapsed or not
   */
  private boolean collapsed;
  /**
   * Indicates whether this row has an explicit default format
   */
  private boolean defaultFormat;
  /**
   * Indicates whether the row record height matches the default font height
   */
  private boolean matchesDefFontHeight;
  /**
   * The (default) xf index for cells on this row
   */
  private int xfIndex;

  /** 
   * The outline level of the row
   */
  private int outlineLevel;

  /** 
   * Is this the icon indicator row of a group?
   */
  private boolean groupStart;

  /**
   * Indicates that the row is default height
   */
  private static final int defaultHeightIndicator = 0xff;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   */
  RowRecord(Record t)
  {
    super(t);

    byte[] data = getRecord().getData();
    rowNumber = IntegerHelper.getInt(data[0], data[1]);
    rowHeight = IntegerHelper.getInt(data[6], data[7]);

    int options = IntegerHelper.getInt(data[12], data[13],
                                       data[14], data[15]);
    outlineLevel = (options & 0x7);
    groupStart = (options & 0x10) != 0;
    collapsed = (options & 0x20) != 0;
    matchesDefFontHeight = (options & 0x40) == 0;
    defaultFormat = (options & 0x80) != 0;
    xfIndex = (options & 0x0fff0000) >> 16;
  }

  /**
   * Interrogates whether this row is of default height
   *
   * @return TRUE if this is set to the default height, FALSE otherwise
   */
  boolean isDefaultHeight()
  {
    return rowHeight == defaultHeightIndicator;
  }

  /**
   * Interrogates this row to see whether it matches the default font height
   *
   * @return TRUE if this matches the default font height, FALSE otherwise
   */
  public boolean matchesDefaultFontHeight()
  {
    return matchesDefFontHeight;
  }

  /**
   * Gets the row number
   *
   * @return the number of this row
   */
  public int getRowNumber()
  {
    return rowNumber;
  }

  /** 
   * Accessor for the row's outline level
   *
   * @return the row's outline level
   */
  public int getOutlineLevel() 
  {
    return outlineLevel;
  }

  /** 
   * Accessor for row's groupStart value
   *
   * @return the row's groupStart value
   */
  public boolean getGroupStart() 
  {
    return groupStart;
  }

  /**
   * Gets the height of the row
   *
   * @return the row height
   */
  public int getRowHeight()
  {
    return rowHeight;
  }

  /**
   * Queries whether the row is collapsed
   *
   * @return the collapsed indicator
   */
  public boolean isCollapsed()
  {
    return collapsed;
  }

  /**
   * Gets the default xf index for this row
   *
   * @return the xf index
   */
  public int getXFIndex()
  {
    return xfIndex;
  }

  /**
   * Queries whether the row has a specific default cell format applied
   *
   * @return the default cell format
   */
  public boolean hasDefaultFormat()
  {
    return defaultFormat;
  }
}


