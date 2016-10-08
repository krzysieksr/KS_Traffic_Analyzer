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

import jxl.Cell;
import jxl.CellFeatures;
import jxl.CellType;
import jxl.biff.FormattingRecords;
import jxl.format.CellFormat;

/**
 * A blank cell value, initialized indirectly from a multiple biff record
 * rather than directly from the binary data
 */
class MulBlankCell implements Cell, CellFeaturesAccessor
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(MulBlankCell.class);

  /**
   * The row containing this blank
   */
  private int row;
  /**
   * The column containing this blank
   */
  private int column;
  /**
   * The raw cell format
   */
  private CellFormat cellFormat;

  /**
   * The index to the XF Record
   */
  private int xfIndex;

  /**
   * A handle to the formatting records
   */
  private FormattingRecords formattingRecords;

  /**
   * A flag to indicate whether this object's formatting things have
   * been initialized
   */
  private boolean initialized;

  /**
   * A handle to the sheet
   */
  private SheetImpl sheet;

  /**
   * The cell features
   */
  private CellFeatures features;


  /**
   * Constructs this cell
   *
   * @param r the zero based row
   * @param c the zero base column
   * @param xfi the xf index
   * @param fr the formatting records
   * @param si the sheet
   */
  public MulBlankCell(int r, int c,
                      int xfi,
                      FormattingRecords fr,
                      SheetImpl si)
  {
    row = r;
    column = c;
    xfIndex = xfi;
    formattingRecords = fr;
    sheet = si;
    initialized = false;
  }

  /**
   * Accessor for the row
   *
   * @return the zero based row
   */
  public final int getRow()
  {
    return row;
  }

  /**
   * Accessor for the column
   *
   * @return the zero based column
   */
  public final int getColumn()
  {
    return column;
  }

  /**
   * Accessor for the contents as a string
   *
   * @return the value as a string
   */
  public String getContents()
  {
    return "";
  }

  /**
   * Accessor for the cell type
   *
   * @return the cell type
   */
  public CellType getType()
  {
    return CellType.EMPTY;
  }

  /**
   * Gets the cell format for this cell
   *
   * @return  the cell format for these cells
   */
  public CellFormat getCellFormat()
  {
    if (!initialized)
    {
      cellFormat = formattingRecords.getXFRecord(xfIndex);
      initialized = true;
    }

    return cellFormat;
  }

  /**
   * Determines whether or not this cell has been hidden
   *
   * @return TRUE if this cell has been hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    ColumnInfoRecord cir = sheet.getColumnInfo(column);

    if (cir != null && cir.getWidth() == 0)
    {
      return true;
    }

    RowRecord rr = sheet.getRowInfo(row);

    if (rr != null && (rr.getRowHeight() == 0 || rr.isCollapsed()))
    {
      return true;
    }

    return false;
  }

  /**
   * Accessor for the cell features
   *
   * @return the cell features or NULL if this cell doesn't have any
   */
  public CellFeatures getCellFeatures()
  {
    return features;
  }

  /**
   * Sets the cell features during the reading process
   *
   * @param cf the cell features
   */
  public void setCellFeatures(CellFeatures cf)
  {
    if (features != null)
    {
      logger.warn("current cell features not null - overwriting");
    }

    features = cf;
  }
}
