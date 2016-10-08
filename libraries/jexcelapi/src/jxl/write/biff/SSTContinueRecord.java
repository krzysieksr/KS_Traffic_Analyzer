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

import java.util.ArrayList;
import java.util.Iterator;

import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * A continuation of a shared string table record.
 */
class SSTContinueRecord extends WritableRecordData
{
  /**
   * The first string is a special case
   */
  private String firstString;
  /**
   * Indicates whether or not we need to include the length information
   * for the first string
   */
  private boolean includeLength;
  /**
   * The length of the first string
   */
  private int firstStringLength;
  /**
   * The list of strings
   */
  private ArrayList strings;
  /**
   * The list of string lengths
   */
  private ArrayList stringLengths;
  /**
   * The binary data
   */
  private byte[] data;
  /**
   * The count of bytes needed so far to contain this record
   */
  private int byteCount;

  /**
   * The maximum amount of bytes available for the SST record
   */
  private static int maxBytes = 8228 - // max length
                                4;  // standard biff record stuff

  /**
   * Constructor
   *
   * @param numRefs the number of string references in the workbook
   * @param s the number of strings
   */
  public SSTContinueRecord()
  {
    super(Type.CONTINUE);

    byteCount = 0;
    strings = new ArrayList(50);
    stringLengths = new ArrayList(50);
  }

  /**
   * Adds the first string to this SST record
   *
   * @param s the string to add
   * @param b include the length information for the first string
   * @return the number of characters not added
   */
  public int setFirstString(String s, boolean b)
  {
    includeLength = b;
    firstStringLength = s.length();

    int bytes = 0;

    if (!includeLength)
    {
      bytes = s.length() * 2 + 1;
    }
    else
    {
      bytes = s.length() * 2 + 3;
    }

    if (bytes <= maxBytes)
    {
      firstString = s;
      byteCount += bytes;
      return 0;
    }

    // Calculate the number of characters we can add
    // The bytes variable will always be an odd number
    int charsAvailable = includeLength ? (maxBytes - 4) / 2 : 
                                         (maxBytes - 2) / 2;

    // Add what part of the string we can
    firstString = s.substring(0, charsAvailable);
    byteCount = maxBytes - 1;

    return s.length() - charsAvailable;
  }

  /**
   * Gets the current offset into this record, excluding the header fields
   *
   * @return the number of bytes after the header field
   */
  public int getOffset()
  {
    return byteCount;
  }

  /**
   * Adds a string to this  record.  It returns the number of string
   * characters not added, due to space constraints.  In the event
   * of this being non-zero, a continue record will be needed
   *
   * @param s the string to add
   * @return the number of characters not added
   */
  public int add(String s)
  {
    int bytes = s.length() * 2 + 3;

    // Must be able to add at least the first character of the string
    // onto the SST
    if (byteCount >= maxBytes - 5)
    {
      return s.length();
    }

    stringLengths.add(new Integer(s.length()));

    if (bytes + byteCount < maxBytes)
    {
      // add the string and return
      strings.add(s);
      byteCount += bytes;
      return 0;
    }

    // Calculate the number of characters we can add
    int bytesLeft = maxBytes - 3 - byteCount;
    int charsAvailable = bytesLeft % 2 == 0 ? bytesLeft / 2 :
                                             (bytesLeft - 1) / 2;

    // Add what part of the string we can
    strings.add(s.substring(0, charsAvailable));
    byteCount += charsAvailable * 2 + 3;

    return s.length() - charsAvailable;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    data = new byte[byteCount];

    int pos = 0;

    // Write out the first string
    if (includeLength)
    {
      IntegerHelper.getTwoBytes(firstStringLength, data, 0);
      data[2] = 0x01;
      pos = 3;
    }
    else
    {
      // Just include the unicode indicator
      data[0] = 0x01;
      pos = 1;
    }

    StringHelper.getUnicodeBytes(firstString, data, pos);
    pos += firstString.length() * 2;

    // Now write out the remainder of the strings
    Iterator i = strings.iterator();
    String s = null;
    int length = 0;
    int count = 0;
    while (i.hasNext())
    {
      s = (String) i.next();
      length = ( (Integer) stringLengths.get(count)).intValue();
      IntegerHelper.getTwoBytes(length, data, pos);
      data[pos+2] = 0x01;
      StringHelper.getUnicodeBytes(s, data, pos+3);
      pos += s.length() * 2 + 3;
      count++;
    }
    
    return data;
  }
}
