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

package jxl.biff;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.format.Format;
import jxl.read.biff.Record;

/**
 * A non-built in format record
 */
public class FormatRecord extends WritableRecordData
  implements DisplayFormat, Format
{
  /**
   * The logger
   */
  public static Logger logger = Logger.getLogger(FormatRecord.class);

  /**
   * Initialized flag
   */
  private boolean initialized;

  /**
   * The raw data
   */
  private byte[] data;

  /**
   * The index code
   */
  private int indexCode;

  /**
   * The formatting string
   */
  private String formatString;

  /**
   * Indicates whether this is a date formatting record
   */
  private boolean date;

  /**
   * Indicates whether this a number formatting record
   */
  private boolean number;

  /**
   * The format object
   */
  private java.text.Format format;

  /**
   * The date strings to look for
   */
  private static String[] dateStrings = new String[]
  {
    "dd",
    "mm",
    "yy",
    "hh",
    "ss",
    "m/",
    "/d"
  };

  // Type to distinguish between biff7 and biff8
  private static class BiffType
  {
  }

  public static final BiffType biff8 = new BiffType();
  public static final BiffType biff7 = new BiffType();

  /**
   * Constructor invoked when copying sheets
   *
   * @param fmt the format string
   * @param refno the index code
   */
  FormatRecord(String fmt, int refno)
  {
    super(Type.FORMAT);
    formatString = fmt;
    indexCode    = refno;
    initialized  = true;
  }

  /**
   * Constructor used by writable formats
   */
  protected FormatRecord()
  {
    super(Type.FORMAT);
    initialized  = false;
  }

  /**
   * Copy constructor - can be invoked by public access
   *
   * @param fr the format to copy
   */
  protected FormatRecord(FormatRecord fr)
  {
    super(Type.FORMAT);
    initialized = false;

    formatString = fr.formatString;
    date = fr.date;
    number = fr.number;
    //    format = (java.text.Format) fr.format.clone();
  }

  /**
   * Constructs this object from the raw data.  Used when reading in a
   * format record
   *
   * @param t the raw data
   * @param ws the workbook settings
   * @param biffType biff type dummy overload
   */
  public FormatRecord(Record t, WorkbookSettings ws, BiffType biffType)
  {
    super(t);

    byte[] data = getRecord().getData();
    indexCode = IntegerHelper.getInt(data[0], data[1]);
    initialized = true;

    if (biffType == biff8)
    {
      int numchars = IntegerHelper.getInt(data[2], data[3]);
      if (data[4] == 0)
      {
        formatString = StringHelper.getString(data, numchars, 5, ws);
      }
      else
      {
        formatString = StringHelper.getUnicodeString(data, numchars, 5);
      }
    }
    else
    {
      int numchars = data[2];
      byte[] chars = new byte[numchars];
      System.arraycopy(data, 3, chars, 0, chars.length);
      formatString = new String(chars);
    }

    date = false;
    number = false;

    // First see if this is a date format
    for (int i = 0 ; i < dateStrings.length; i++)
    {
      String dateString = dateStrings[i];
      if (formatString.indexOf(dateString) != -1 || 
          formatString.indexOf(dateString.toUpperCase()) != -1)
      {
        date = true;
        break;
      }
    }

    // See if this is number format - look for the # or 0 characters
    if (!date)
    {
      if (formatString.indexOf('#') != -1 ||
          formatString.indexOf('0') != -1 )
      {
        number = true;
      }
    }
  }

  /**
   * Used to get the data when writing out the format record
   *
   * @return the raw data
   */
  public byte[] getData()
  {
    data = new byte[formatString.length() * 2 + 3 + 2];

    IntegerHelper.getTwoBytes(indexCode, data, 0);
    IntegerHelper.getTwoBytes(formatString.length(), data, 2);
    data[4] = (byte) 1; // unicode indicator
    StringHelper.getUnicodeBytes(formatString, data, 5);

    return data;
  }

  /**
   * Gets the format index of this record
   *
   * @return the format index of this record
   */
  public int getFormatIndex()
  {
    return indexCode;
  }

  /**
   * Accessor to see whether this object is initialized or not.
   *
   * @return TRUE if this font record has been initialized, FALSE otherwise
   */
  public boolean isInitialized()
  {
    return initialized;
  }

  /**
   * Sets the index of this record.  Called from the FormattingRecords
   * object
   *
   * @param pos the position of this font in the workbooks font list
   */

  public void initialize(int pos)
  {
    indexCode = pos;
    initialized = true;
  }

  /**
   * Replaces all instances of search with replace in the input.  Used for
   * replacing microsoft number formatting characters with java equivalents
   *
   * @param input the format string
   * @param search the Excel character to be replaced
   * @param replace the java equivalent
   * @return the input string with the specified substring replaced
   */
  protected final String replace(String input, String search, String replace)
  {
    String fmtstr = input;
    int pos = fmtstr.indexOf(search);
    while (pos != -1)
    {
      StringBuffer tmp = new StringBuffer(fmtstr.substring(0, pos));
      tmp.append(replace);
      tmp.append(fmtstr.substring(pos + search.length()));
      fmtstr = tmp.toString();
      pos = fmtstr.indexOf(search);
    }
    return fmtstr;
  }

  /**
   * Called by the immediate subclass to set the string
   * once the Java-Excel replacements have been done
   *
   * @param s the format string
   */
  protected final void setFormatString(String s)
  {
    formatString = s;
  }

  /**
   * Sees if this format is a date format
   *
   * @return TRUE if this format is a date
   */
  public final  boolean isDate()
  {
    return date;
  }

  /**
   * Sees if this format is a number format
   *
   * @return TRUE if this format is a number
   */
  public final boolean isNumber()
  {
    return number;
  }

  /**
   * Gets the java equivalent number format for the formatString
   *
   * @return The java equivalent of the number format for this object
   */
  public final NumberFormat getNumberFormat()
  {
    if (format != null && format instanceof NumberFormat)
    {
      return (NumberFormat) format;
    }

    try
    {
      String fs = formatString;

      // Replace the Excel formatting characters with java equivalents
      fs = replace(fs, "E+", "E");
      fs = replace(fs, "_)", "");
      fs = replace(fs, "_", "");
      fs = replace(fs, "[Red]", "");
      fs = replace(fs, "\\", "");

      format = new DecimalFormat(fs);
    }
    catch (IllegalArgumentException e)
    {
      // Something went wrong with the date format - fail silently
      // and return a default value
      format = new DecimalFormat("#.###");
    }

    return (NumberFormat) format;
  }

  /**
   * Gets the java equivalent date format for the formatString
   *
   * @return The java equivalent of the date format for this object
   */
  public final DateFormat getDateFormat()
  {
    if (format != null && format instanceof DateFormat)
    {
      return (DateFormat) format;
    }

    String fmt = formatString;

    // Replace the AM/PM indicator with an a
    int pos = fmt.indexOf("AM/PM");
    while (pos != -1)
    {
      StringBuffer sb = new StringBuffer(fmt.substring(0, pos));
      sb.append('a');
      sb.append(fmt.substring(pos + 5));
      fmt = sb.toString();
      pos = fmt.indexOf("AM/PM");
    }

    // Replace ss.0 with ss.SSS (necessary to always specify milliseconds
    // because of NT)
    pos = fmt.indexOf("ss.0");
    while (pos != -1)
    {
      StringBuffer sb = new StringBuffer(fmt.substring(0, pos));
      sb.append("ss.SSS");

      // Keep going until we run out of zeros
      pos += 4;
      while (pos < fmt.length() && fmt.charAt(pos) == '0')
      {
        pos++;
      }

      sb.append(fmt.substring(pos));
      fmt = sb.toString();
      pos = fmt.indexOf("ss.0");
    }


    // Filter out the backslashes
    StringBuffer sb = new StringBuffer();
    for (int i = 0; i < fmt.length(); i++)
    {
      if (fmt.charAt(i) != '\\')
      {
        sb.append(fmt.charAt(i));
      }
    }

    fmt = sb.toString();

    // If the date format starts with anything inside square brackets then 
    // filter tham out
    if (fmt.charAt(0) == '[')
    {
      int end = fmt.indexOf(']');
      if (end != -1)
      {
        fmt = fmt.substring(end+1);
      }
    }
    
    // Get rid of some spurious characters that can creep in
    fmt = replace(fmt, ";@", "");

    // We need to convert the month indicator m, to upper case when we
    // are dealing with dates
    char[] formatBytes = fmt.toCharArray();

    for (int i = 0; i < formatBytes.length; i++)
    {
      if (formatBytes[i] == 'm')
      {
        // Firstly, see if the preceding character is also an m.  If so,
        // copy that
        if (i > 0 && (formatBytes[i - 1] == 'm' || formatBytes[i - 1] == 'M'))
        {
          formatBytes[i] = formatBytes[i - 1];
        }
        else
        {
          // There is no easy way out.  We have to deduce whether this an
          // minute or a month?  See which is closest out of the
          // letters H d s or y
          // First, h
          int minuteDist = Integer.MAX_VALUE;
          for (int j = i - 1; j > 0; j--)
          {
            if (formatBytes[j] == 'h')
            {
              minuteDist = i - j;
              break;
            }
          }

          for (int j = i + 1; j < formatBytes.length; j++)
          {
            if (formatBytes[j] == 'h')
            {
              minuteDist = Math.min(minuteDist, j - i);
              break;
            }
          }

          for (int j = i - 1; j > 0; j--)
          {
            if (formatBytes[j] == 'H')
            {
              minuteDist = i - j;
              break;
            }
          }

          for (int j = i + 1; j < formatBytes.length; j++)
          {
            if (formatBytes[j] == 'H')
            {
              minuteDist = Math.min(minuteDist, j - i);
              break;
            }
          }

          // Now repeat for s
          for (int j = i - 1; j > 0; j--)
          {
            if (formatBytes[j] == 's')
            {
              minuteDist = Math.min(minuteDist, i - j);
              break;
            }
          }
          for (int j = i + 1; j < formatBytes.length; j++)
          {
            if (formatBytes[j] == 's')
            {
              minuteDist = Math.min(minuteDist, j - i);
              break;
            }
          }
          // We now have the distance of the closest character which could
          // indicate the the m refers to a minute
          // Repeat for d and y
          int monthDist = Integer.MAX_VALUE;
          for (int j = i - 1; j > 0; j--)
          {
            if (formatBytes[j] == 'd')
            {
              monthDist = i - j;
              break;
            }
          }

          for (int j = i + 1; j < formatBytes.length; j++)
          {
            if (formatBytes[j] == 'd')
            {
              monthDist = Math.min(monthDist, j - i);
              break;
            }
          }
          // Now repeat for y
          for (int j = i - 1; j > 0; j--)
          {
            if (formatBytes[j] == 'y')
            {
              monthDist = Math.min(monthDist, i - j);
              break;
            }
          }
          for (int j = i + 1; j < formatBytes.length; j++)
          {
            if (formatBytes[j] == 'y')
            {
              monthDist = Math.min(monthDist, j - i);
              break;
            }
          }

          if (monthDist < minuteDist)
          {
            // The month indicator is closer, so convert to a capital M
            formatBytes[i] = Character.toUpperCase(formatBytes[i]);
          }
          else if ((monthDist == minuteDist) &&
                   (monthDist != Integer.MAX_VALUE))
          {
            // They are equidistant.  As a tie-breaker, take the formatting
            // character which precedes the m
            char ind = formatBytes[i - monthDist];
            if (ind == 'y' || ind == 'd')
            {
              // The preceding item indicates a month measure, so convert
              formatBytes[i] = Character.toUpperCase(formatBytes[i]);
            }
          }
        }
      }
    }

    try
    {
      this.format = new SimpleDateFormat(new String(formatBytes));
    }
    catch (IllegalArgumentException e)
    {
      // There was a spurious character - fail silently
      this.format = new SimpleDateFormat("dd MM yyyy hh:mm:ss");
    }
    return (DateFormat) this.format;
  }

  /**
   * Gets the index code, for use as a hash value
   *
   * @return the ifmt code for this cell
   */
  public int getIndexCode()
  {
    return indexCode;
  }

  /**
   * Gets the formatting string.
   *
   * @return the excel format string
   */
  public String getFormatString()
  {
    return formatString;
  }

  /**
   * Indicates whether this formula is a built in
   *
   * @return FALSE
   */
  public boolean isBuiltIn()
  {
    return false;
  }

  /**
   * Standard hash code method
   * @return the hash code value for this object
   */
  public int hashCode()
  {
    return formatString.hashCode();
  }

  /**
   * Standard equals method.  This compares the contents of two
   * format records, and not their indexCodes, which are ignored
   *
   * @param o the object to compare
   * @return TRUE if the two objects are equal, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof FormatRecord))
    {
      return false;
    }

    FormatRecord fr = (FormatRecord) o;

    // Initialized format comparison
    if (initialized && fr.initialized)
    {
      // Must be either a number or a date format
      if (date != fr.date ||
          number != fr.number)
      {
        return false;
      }

      return formatString.equals(fr.formatString);
    }

    // Uninitialized format comparison
    return formatString.equals(fr.formatString);
  }
}






