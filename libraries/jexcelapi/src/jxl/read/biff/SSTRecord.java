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

import jxl.common.Assert;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.StringHelper;

/**
 * Holds all the strings in the shared string table
 */
class SSTRecord extends RecordData
{
  /**
   * The total number of strings in this table
   */
  private int totalStrings;
  /**
   * The number of unique strings
   */
  private int uniqueStrings;
  /**
   * The shared strings
   */
  private String[] strings;
  /**
   * The array of continuation breaks
   */
  private int[] continuationBreaks;

  /**
   * A holder for a byte array
   */
  private static class ByteArrayHolder
  {
    /**
     * the byte holder
     */
    public byte[] bytes;
  }

  /**
   * A holder for a boolean
   */
  private static class BooleanHolder
  {
    /**
     * the holder holder
     */
    public boolean value;
  }

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param continuations the continuations
   * @param ws the workbook settings
   */
  public SSTRecord(Record t, Record[] continuations, WorkbookSettings ws)
  {
    super(t);

    // If a continue record appears in the middle of
    // a string, then the encoding character is repeated

    // Concatenate everything into one big bugger of a byte array
    int totalRecordLength = 0;

    for (int i = 0; i < continuations.length; i++)
    {
      totalRecordLength += continuations[i].getLength();
    }
    totalRecordLength += getRecord().getLength();

    byte[] data = new byte[totalRecordLength];

    // First the original data gets put in
    int pos = 0;
    System.arraycopy(getRecord().getData(), 0,
                     data, 0, getRecord().getLength());
    pos += getRecord().getLength();

    // Now copy in everything else.
    continuationBreaks = new int[continuations.length];
    Record r = null;
    for (int i = 0; i < continuations.length; i++)
    {
      r = continuations[i];
      System.arraycopy(r.getData(), 0,
                       data, pos,
                       r.getLength());
      continuationBreaks[i] = pos;
      pos += r.getLength();
    }

    totalStrings = IntegerHelper.getInt(data[0], data[1],
                                        data[2], data[3]);
    uniqueStrings = IntegerHelper.getInt(data[4], data[5],
                                         data[6], data[7]);

    strings = new String[uniqueStrings];
    readStrings(data, 8, ws);
  }

  /**
   * Reads in all the strings from the raw data
   *
   * @param data the raw data
   * @param offset the offset
   * @param ws the workbook settings
   */
  private void readStrings(byte[] data, int offset, WorkbookSettings ws)
  {
    int pos = offset;
    int numChars;
    byte optionFlags;
    String s = null;
    boolean asciiEncoding = false;
    boolean richString = false;
    boolean extendedString = false;
    int formattingRuns = 0;
    int extendedRunLength = 0;

    for (int i = 0; i < uniqueStrings; i++)
    {
      // Read in the number of characters
      numChars = IntegerHelper.getInt(data[pos], data[pos + 1]);
      pos += 2;
      optionFlags = data[pos];
      pos++;

      // See if it is an extended string
      extendedString = ((optionFlags & 0x04) != 0);

      // See if string contains formatting information
      richString = ((optionFlags & 0x08) != 0);

      if (richString)
      {
        // Read in the crun
        formattingRuns = IntegerHelper.getInt(data[pos], data[pos + 1]);
        pos += 2;
      }

      if (extendedString)
      {
        // Read in cchExtRst
        extendedRunLength = IntegerHelper.getInt
          (data[pos], data[pos + 1], data[pos + 2], data[pos + 3]);
        pos += 4;
      }

      // See if string is ASCII (compressed) or unicode
      asciiEncoding = ((optionFlags & 0x01) == 0);

      ByteArrayHolder bah = new ByteArrayHolder();
      BooleanHolder   bh = new BooleanHolder();
      bh.value = asciiEncoding;
      pos += getChars(data, bah, pos, bh, numChars);
      asciiEncoding = bh.value;

      if (asciiEncoding)
      {
        s = StringHelper.getString(bah.bytes, numChars, 0, ws);
      }
      else
      {
        s = StringHelper.getUnicodeString(bah.bytes, numChars, 0);
      }

      strings[i] = s;

      // For rich strings, skip over the formatting runs
      if (richString)
      {
        pos += 4 * formattingRuns;
      }

      // For extended strings, skip over the extended string data
      if (extendedString)
      {
        pos += extendedRunLength;
      }

      if (pos > data.length)
      {
        Assert.verify(false, "pos exceeds record length");
      }
    }
  }

  /**
   * Gets the chars in the ascii array, taking into account continuation
   * breaks
   *
   * @param source the original source
   * @param bah holder for the new byte array
   * @param pos the current position in the source
   * @param ascii holder for a return ascii flag
   * @param numChars the number of chars in the string
   * @return the number of bytes read from the source
   */
  private int getChars(byte[] source,
                       ByteArrayHolder bah,
                       int pos,
                       BooleanHolder ascii,
                       int numChars)
  {
    int i = 0;
    boolean spansBreak = false;

    if (ascii.value)
    {
      bah.bytes = new byte[numChars];
    }
    else
    {
      bah.bytes = new byte[numChars * 2];
    }

    while (i < continuationBreaks.length && !spansBreak)
    {
      spansBreak = pos <= continuationBreaks[i] &&
                   (pos + bah.bytes.length > continuationBreaks[i]);

      if (!spansBreak)
      {
        i++;
      }
    }

    // If it doesn't span a break simply do an array copy into the
    // destination array and finish
    if (!spansBreak)
    {
      System.arraycopy(source, pos, bah.bytes, 0, bah.bytes.length);
      return bah.bytes.length;
    }

    // Copy the portion before the break pos into the array
    int breakpos = continuationBreaks[i];
    System.arraycopy(source, pos, bah.bytes, 0, breakpos - pos);

    int bytesRead = breakpos - pos;
    int charsRead;
    if (ascii.value)
    {
      charsRead = bytesRead;
    }
    else
    {
      charsRead = bytesRead / 2;
    }

    bytesRead += getContinuedString(source,
                                    bah,
                                    bytesRead,
                                    i,
                                    ascii,
                                    numChars - charsRead);
    return bytesRead;
  }

  /**
   * Gets the rest of the string after a continuation break
   *
   * @param source the original bytes
   * @param bah the holder for the new bytes
   * @param destPos the des pos
   * @param contBreakIndex the index of the continuation break
   * @param ascii the ascii flag holder
   * @param charsLeft the number of chars left in the array
   * @return the number of bytes read in the continued string
   */
  private int getContinuedString(byte[] source,
                                 ByteArrayHolder bah,
                                 int destPos,
                                 int contBreakIndex,
                                 BooleanHolder ascii,
                                 int charsLeft)
  {
    int breakpos = continuationBreaks[contBreakIndex];
    int bytesRead = 0;

    while (charsLeft > 0)
    {
      Assert.verify(contBreakIndex < continuationBreaks.length,
                    "continuation break index");

      if (ascii.value && source[breakpos] == 0)
      {
        // The string is consistently ascii throughout

        int length = contBreakIndex == continuationBreaks.length - 1 ?
          charsLeft :
          Math.min
            (charsLeft,
             continuationBreaks[contBreakIndex + 1] - breakpos - 1);

        System.arraycopy(source,
                         breakpos + 1,
                         bah.bytes,
                         destPos,
                         length);
        destPos   += length;
        bytesRead += length + 1;
        charsLeft -= length;
        ascii.value = true;
      }
      else if (!ascii.value && source[breakpos] != 0)
      {
        // The string is Unicode throughout

        int length = contBreakIndex == continuationBreaks.length - 1 ?
          charsLeft * 2 :
          Math.min
            (charsLeft * 2,
             continuationBreaks[contBreakIndex + 1] - breakpos - 1);

        // It looks like the string continues as Unicode too.  That's handy
        System.arraycopy(source,
                         breakpos + 1,
                         bah.bytes,
                         destPos,
                         length);

        destPos   += length;
        bytesRead += length + 1;
        charsLeft -= length / 2;
        ascii.value = false;
      }
      else if (!ascii.value && source[breakpos] == 0)
      {
        // Bummer - the string starts off as Unicode, but after the
        // continuation it is in straightforward ASCII encoding
        int chars = contBreakIndex == continuationBreaks.length - 1 ?
          charsLeft:
          Math.min
            (charsLeft,
             continuationBreaks[contBreakIndex + 1] - breakpos - 1);

        for (int j = 0; j < chars; j++)
        {
          bah.bytes[destPos] = source[breakpos + j + 1];
          destPos += 2;
        }

        bytesRead += chars + 1;
        charsLeft -= chars;
        ascii.value = false;
      }
      else
      {
        // Double Bummer - the string starts off as ASCII, but after the
        // continuation it is in Unicode.  This impacts the allocated array

        // Reallocate what we have of the byte array so that it is all
        // Unicode
        byte[] oldBytes = bah.bytes;
        bah.bytes = new byte[destPos * 2 + charsLeft * 2];
        for (int j = 0; j < destPos; j++)
        {
          bah.bytes[j * 2] = oldBytes[j];
        }

        destPos = destPos * 2;

        int length = contBreakIndex == continuationBreaks.length - 1 ?
          charsLeft * 2 :
          Math.min
            (charsLeft * 2,
             continuationBreaks[contBreakIndex + 1] - breakpos - 1);

        System.arraycopy(source,
                         breakpos + 1,
                         bah.bytes,
                         destPos,
                         length);

        destPos   += length;
        bytesRead += length + 1;
        charsLeft -= length / 2;
        ascii.value = false;
      }

      contBreakIndex++;

      if (contBreakIndex < continuationBreaks.length)
      {
        breakpos = continuationBreaks[contBreakIndex];
      }
    }

    return bytesRead;
  }

  /**
   * Gets the string at the specified position
   *
   * @param index the index of the string to return
   * @return the strings
   */
  public String getString(int index)
  {
    Assert.verify(index < uniqueStrings);
    return strings[index];
  }
}


