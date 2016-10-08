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
import jxl.biff.FormattingRecords;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.XFRecord;
import jxl.format.CellFormat;

/**
 * Abstract class for all records which actually contain cell values
 */
public abstract class CellValue extends RecordData
  implements Cell, CellFeaturesAccessor
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CellValue.class);

  /**
   * The row number of this cell record
   */
  private int row;

  /**
   * The column number of this cell record
   */
  private int column;

  /**
   * The XF index
   */
  private int xfIndex;

  /**
   * A handle to the formatting records, so that we can
   * retrieve the formatting information
   */
  private FormattingRecords formattingRecords;

  /**
   * A lazy initialize flag for the cell format
   */
  private boolean initialized;

  /**
   * The cell format
   */
  private XFRecord format;

  /**
   * A handle back to the sheet
   */
  private SheetImpl sheet;

  /**
   * The cell features
   */
  private CellFeatures features;

  /**
   * Constructs this object from the raw cell data
   *
   * @param t the raw cell data
   * @param fr the formatting records
   * @param si the sheet containing this cell
   */
  protected CellValue(Record t, FormattingRecords fr, SheetImpl si)
  {
    super(t);
    byte[] data = getRecord().getData();
    row     = IntegerHelper.getInt(data[0], data[1]);
    column  = IntegerHelper.getInt(data[2], data[3]);
    xfIndex = IntegerHelper.getInt(data[4], data[5]);
    sheet = si;
    formattingRecords = fr;
    initialized = false;
  }

  /**
   * Interface method which returns the row number of this cell
   *
   * @return the zero base row number
   */
  public final int getRow()
  {
    return row;
  }

  /**
   * Interface method which returns the column number of this cell
   *
   * @return the zero based column number
   */
  public final int getColumn()
  {
    return column;
  }

  /**
   * Gets the XFRecord corresponding to the index number.  Used when
   * copying a spreadsheet
   *
   * @return the xf index for this cell
   */
  public final int getXFIndex()
  {
    return xfIndex;
  }

  /**
   * Gets the CellFormat object for this cell.  Used by the WritableWorkbook
   * API
   *
   * @return the CellFormat used for this cell
   */
  public CellFormat getCellFormat()
  {
    if (!initialized)
    {
      format = formattingRecords.getXFRecord(xfIndex);
      initialized = true;
    }

    return format;
  }

  /**
   * Determines whether or not this cell has been hidden
   *
   * @return TRUE if this cell has been hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    ColumnInfoRecord cir = sheet.getColumnInfo(column);

    if (cir != null && (cir.getWidth() == 0 || cir.getHidden()))
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
   * Accessor for the sheet
   *
   * @return the sheet
   */
  protected SheetImpl getSheet()
  {
    return sheet;
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

