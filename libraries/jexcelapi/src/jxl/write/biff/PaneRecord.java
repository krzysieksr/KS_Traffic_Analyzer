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

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Contains the window attributes for a worksheet
 */
class PaneRecord extends WritableRecordData
{
  /**
   * The number of rows visible in the top left pane
   */
  private int rowsVisible;
  /**
   * The number of columns visible in the top left pane
   */
  private int columnsVisible;

  /**
   * The pane codes
   */
  private final static int topLeftPane = 0x3;
  private final static int bottomLeftPane = 0x2;
  private final static int topRightPane = 0x1;
  private final static int bottomRightPane = 0x0;

  /**
   * Code

  /**
   * Constructor
   */
  public PaneRecord(int cols, int rows)
  {
    super(Type.PANE);

    rowsVisible = rows;
    columnsVisible = cols;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[10];

    // The x position
    IntegerHelper.getTwoBytes(columnsVisible, data, 0);

    // The y position
    IntegerHelper.getTwoBytes(rowsVisible, data, 2);

    // The top row visible in the bottom pane
    if (rowsVisible > 0)
    {
      IntegerHelper.getTwoBytes(rowsVisible, data, 4);
    }

    // The left most column visible in the right pane
    if (columnsVisible > 0)
    {
      IntegerHelper.getTwoBytes(columnsVisible, data, 6);
    }

    // The active pane
    int activePane = topLeftPane;

    if (rowsVisible > 0 && columnsVisible == 0)
    {
      activePane = bottomLeftPane;
    }
    else if (rowsVisible == 0 && columnsVisible > 0)
    {
      activePane = topRightPane;
    }
    else if (rowsVisible > 0 && columnsVisible > 0)
    {
      activePane = bottomRightPane;
    }
    // always present
    IntegerHelper.getTwoBytes(activePane, data, 8);

    return data;
  }
}
