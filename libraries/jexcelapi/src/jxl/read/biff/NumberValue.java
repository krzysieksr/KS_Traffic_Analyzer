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

import java.text.DecimalFormat;
import java.text.NumberFormat;

import jxl.CellFeatures;
import jxl.CellType;
import jxl.NumberCell;
import jxl.biff.FormattingRecords;
import jxl.format.CellFormat;

/**
 * A numerical cell value, initialized indirectly from a multiple biff record
 * rather than directly from the binary data
 */
class NumberValue implements NumberCell, CellFeaturesAccessor
{
  /**
   * The row containing this number
   */
  private int row;
  /**
   * The column containing this number
   */
  private int column;
  /**
   * The value of this number
   */
  private double value;

  /**
   * The cell format
   */
  private NumberFormat format;

  /**
   * The raw cell format
   */
  private CellFormat cellFormat;

  /**
   * The cell features
   */
  private CellFeatures features;

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
   * The format in which to return this number as a string
   */
  private static DecimalFormat defaultFormat = new DecimalFormat("#.###");

  /**
   * Constructs this number
   *
   * @param r the zero based row
   * @param c the zero base column
   * @param val the value
   * @param xfi the xf index
   * @param fr the formatting records
   * @param si the sheet
   */
  public NumberValue(int r, int c, double val,
                     int xfi,
                     FormattingRecords fr,
                     SheetImpl si)
  {
    row = r;
    column = c;
    value = val;
    format = defaultFormat;
    xfIndex = xfi;
    formattingRecords = fr;
    sheet = si;
    initialized = false;
  }

  /**
   * Sets the format for the number based on the Excel spreadsheets' format.
   * This is called from SheetImpl when it has been definitely established
   * that this cell is a number and not a date
   *
   * @param f the format
   */
  final void setNumberFormat(NumberFormat f)
  {
    if (f != null)
    {
      format = f;
    }
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
   * Accessor for the value
   *
   * @return the value
   */
  public double getValue()
  {
    return value;
  }

  /**
   * Accessor for the contents as a string
   *
   * @return the value as a string
   */
  public String getContents()
  {
    return format.format(value);
  }

  /**
   * Accessor for the cell type
   *
   * @return the cell type
   */
  public CellType getType()
  {
    return CellType.NUMBER;
  }

  /**
   * Gets the cell format
   *
   * @return the cell format
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
   * Gets the NumberFormat used to format this cell.  This is the java
   * equivalent of the Excel format
   *
   * @return the NumberFormat used to format the cell
   */
  public NumberFormat getNumberFormat()
  {
    return format;
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
   * Sets the cell features
   *
   * @param cf the cell features
   */
  public void setCellFeatures(CellFeatures cf)
  {
    features = cf;
  }

}



