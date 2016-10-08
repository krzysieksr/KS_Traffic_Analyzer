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

import java.text.DecimalFormat;

import jxl.biff.DisplayFormat;
import jxl.write.biff.NumberFormatRecord;

/**
 * A custom user defined number format, which may be instantiated within user
 * applications in order to present numerical values to the appropriate level
 * of accuracy.
 * The string format used to create a number format adheres to the standard
 * java specification, and JExcelAPI makes the necessary modifications so
 * that it is rendered in Excel as the nearest possible equivalent.
 * Once created, this may be used within a CellFormat object, which in turn
 * is a parameter passed to the constructor of the Number cell
 */
public class NumberFormat extends NumberFormatRecord implements DisplayFormat
{
  /**
   * Pass in to the constructor to bypass the format validation
   */
  public static final NonValidatingFormat COMPLEX_FORMAT =
    new NumberFormatRecord.NonValidatingFormat();

  // Some format strings

  /**
   * Constant format string for the Euro currency symbol where it precedes
   * the format
   */
  public static final String CURRENCY_EURO_PREFIX = "[$€-2]";

  /**
   * Constant format string for the Euro currency symbol where it precedes
   * the format
   */
  public static final String CURRENCY_EURO_SUFFIX = "[$€-1]";

  /**
   * Constant format string for the UK pound sign
   */
  public static final String CURRENCY_POUND = "£";

  /**
   * Constant format string for the Japanese Yen sign
   */
  public static final String CURRENCY_JAPANESE_YEN = "[$¥-411]";

  /**
   * Constant format string for the US Dollar sign
   */
  public static final String CURRENCY_DOLLAR = "[$$-409]";

  /**
   * Constant format string for three digit fractions
   */
  public static final String FRACTION_THREE_DIGITS = "???/???";

  /**
   * Constant format string for fractions as halves
   */
  public static final String FRACTION_HALVES = "?/2";

  /**
   * Constant format string for fractions as quarter
   */
  public static final String FRACTION_QUARTERS = "?/4";

  /**
   * Constant format string for fractions as eighths
   */
  public static final String FRACTIONS_EIGHTHS = "?/8";

  /**
   * Constant format string for fractions as sixteenths
   */
  public static final String FRACTION_SIXTEENTHS = "?/16";

  /**
   * Constant format string for fractions as tenths
   */
  public static final String FRACTION_TENTHS = "?/10";

  /**
   * Constant format string for fractions as hundredths
   */
  public static final String FRACTION_HUNDREDTHS = "?/100";

  /**
   * Constructor, taking in the Java compliant number format
   *
   * @param format the format string
   */
  public NumberFormat(String format)
  {
    super(format);

    // Verify that the format is valid
    DecimalFormat df = new DecimalFormat(format);
  }

  /**
   * Constructor, taking in the non-Java compliant number format.  This
   * may be used for currencies and more complex custom formats, which
   * will not be subject to the standard validation rules.
   * As there is no validation, there is a resultant risk that the
   * generated Excel file will be corrupt
   *
   * USE THIS CONSTRUCTOR ONLY IF YOU ARE CERTAIN THAT THE NUMBER FORMAT
   * YOU ARE USING IS EXCEL COMPLIANT
   *
   * @param format the format string
   * @param dummy dummy parameter
   */
  public NumberFormat(String format, NonValidatingFormat dummy)
  {
    super(format, dummy);
  }
}
