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

package jxl.write;

import java.text.SimpleDateFormat;

import jxl.biff.DisplayFormat;
import jxl.write.biff.DateFormatRecord;

/**
 * A custom user defined number format which may be instantiated within user
 * applications in order to present date and time values in the  appropriate
 * format.
 * The string format used to create a DateFormat adheres to the standard
 * java specification, and JExcelApi makes the necessary modifications so
 * that it is rendered as its nearest equivalent in Excel.
 * Once created, this may be used within a CellFormat object, which in turn
 * is a parameter passed to the constructor of the DateTime cell
 */
public class DateFormat extends DateFormatRecord implements DisplayFormat
{
  /**
   * Constructor. The date format that is passed should comply to the standard
   * Java date formatting conventions
   *
   * @param format the date format
   */
  public DateFormat(String format)
  {
    super(format);

    // Verify that the format is valid
    SimpleDateFormat df = new SimpleDateFormat(format);
  }
}
