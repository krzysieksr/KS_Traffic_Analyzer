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

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.CellFeatures;
import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.biff.FormattingRecords;
import jxl.format.CellFormat;

/**
 * A date which is stored in the cell
 */
class DateRecord implements DateCell, CellFeaturesAccessor
{
  /**
   * The logger
   */
  private static Logger logger  = Logger.getLogger(DateRecord.class);

  /**
   * The date represented within this cell
   */
  private Date date;
  /**
   * The row number of this cell record
   */
  private int row;
  /**
   * The column number of this cell record
   */
  private int column;

  /**
   * Indicates whether this is a full date, or merely a time
   */
  private boolean time;

  /**
   * The format to use when displaying this cell's contents as a string
   */
  private DateFormat format;

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
   * A handle to the sheet
   */
  private SheetImpl sheet;

  /**
   * The cell features
   */
  private CellFeatures features;


  /**
   * A flag to indicate whether this objects formatting things have
   * been initialized
   */
  private boolean initialized;

  // The default formats used when returning the date as a string
  private static final SimpleDateFormat dateFormat =
   new  SimpleDateFormat("dd MMM yyyy");

  private static final SimpleDateFormat timeFormat =
    new SimpleDateFormat("HH:mm:ss");

  // The number of days between 1 Jan 1900 and 1 March 1900. Excel thinks
  // the day before this was 29th Feb 1900, but it was 28th Feb 1900.
  // I guess the programmers thought nobody would notice that they
  // couldn't be bothered to program this dating anomaly properly
  private static final int nonLeapDay = 61;

  private static final TimeZone gmtZone = TimeZone.getTimeZone("GMT");

  // The number of days between 01 Jan 1900 and 01 Jan 1970 - this gives
  // the UTC offset
  private static final int utcOffsetDays = 25569;

  // The number of days between 01 Jan 1904 and 01 Jan 1970 - this gives
  // the UTC offset using the 1904 date system
  private static final int utcOffsetDays1904 = 24107;

  // The number of milliseconds in  a day
  private static final long secondsInADay = 24 * 60 * 60;
  private static final long msInASecond = 1000;
  private static final long msInADay = secondsInADay * msInASecond;
  

  /**
   * Constructs this object from the raw data
   *
   * @param num the numerical representation of this
   * @param xfi the java equivalent of the excel date format
   * @param fr  the formatting records
   * @param nf  flag indicating whether we are using the 1904 date system
   * @param si  the sheet
   */
  public DateRecord(NumberCell num,
                    int xfi, FormattingRecords fr,
                    boolean nf, SheetImpl si)
  {
    row = num.getRow();
    column = num.getColumn();
    xfIndex = xfi;
    formattingRecords = fr;
    sheet = si;
    initialized = false;

    format = formattingRecords.getDateFormat(xfIndex);

    // This value represents the number of days since 01 Jan 1900
    double numValue = num.getValue();

    if (Math.abs(numValue) < 1)
    {
      if (format == null)
      {
        format = timeFormat;
      }
      time = true;
    }
    else
    {
      if (format == null)
      {
        format = dateFormat;
      }
      time = false;
    }

    // Work round a bug in excel.  Excel seems to think there is a date
    // called the 29th Feb, 1900 - but in actual fact this was not a leap year.
    // Therefore for values less than 61 in the 1900 date system,
    // add one to the numeric value
    if (!nf && !time && numValue < nonLeapDay)
    {
      numValue += 1;
    }

    // Get rid of any timezone adjustments - we are not interested
    // in automatic adjustments
    format.setTimeZone(gmtZone);

    // Convert this to the number of days since 01 Jan 1970
    int offsetDays = nf ? utcOffsetDays1904 : utcOffsetDays;
    double utcDays = numValue - offsetDays;

    // Convert this into utc by multiplying by the number of milliseconds
    // in a day.  Use the round function prior to ms conversion due
    // to a rounding feature of Excel (contributed by Jurgen
    long utcValue = Math.round(utcDays * secondsInADay) * msInASecond;

    date = new Date(utcValue);
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
   * Gets the date
   *
   * @return the date
   */
  public Date getDate()
  {
    return date;
  }

  /**
   * Gets the cell contents as a string.  This method will use the java
   * equivalent of the excel formatting string
   *
   * @return the label
   */
  public String getContents()
  {
    return format.format(date);
  }

  /**
   * Accessor for the cell type
   *
   * @return the cell type
   */
  public CellType getType()
  {
    return CellType.DATE;
  }

  /**
   * Indicates whether the date value contained in this cell refers to a date,
   * or merely a time
   *
   * @return TRUE if the value refers to a time
   */
  public boolean isTime()
  {
    return time;
  }

  /**
   * Gets the DateFormat used to format the cell.  This will normally be
   * the format specified in the excel spreadsheet, but in the event of any
   * difficulty parsing this, it will revert to the default date/time format.
   *
   * @return the DateFormat object used to format the date in the original
   * excel cell
   */
  public DateFormat getDateFormat()
  {
    Assert.verify(format != null);

    return format;
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
   * Accessor for the sheet
   *
   * @return  the containing sheet
   */
  protected final SheetImpl getSheet()
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
   * Sets the cell features
   *
   * @param cf the cell features
   */
  public void setCellFeatures(CellFeatures cf)
  {
    features = cf;
  }

}





