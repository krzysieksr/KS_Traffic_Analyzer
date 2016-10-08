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

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * The list of available shared strings.  This class contains
 * the labels used for the entire spreadsheet
 */
class SharedStrings
{
  /**
   * All the strings in the spreadsheet, keyed on the string itself
   */
  private HashMap strings;

  /**
   * Contains the same strings, held in a list
   */
  private ArrayList stringList;

  /**
   * The total occurrence of strings in the workbook
   */
  private int totalOccurrences;

  /**
   * Constructor
   */
  public SharedStrings()
  {
    strings = new HashMap(100);
    stringList = new ArrayList(100);
    totalOccurrences = 0;
  }

  /**
   * Gets the index for the string passed in.  If the string is already
   * present, then returns the index of that string, otherwise
   * creates a new key-index mapping
   *
   * @param s the string whose index we want
   * @return the index of the string
   */
  public int getIndex(String s)
  {
    Integer i = (Integer) strings.get(s);

    if (i == null)
    {
      i = new Integer(strings.size());
      strings.put(s, i);
      stringList.add(s);
    }

    totalOccurrences++;

    return i.intValue();
  }

  /**
   * Gets the string at the specified index
   *
   * @param i the index of the string
   * @return the string at the specified index
   */
  public String get(int i)
  {
    return (String) stringList.get(i);
  }

  /**
   * Writes out the shared string table
   *
   * @param outputFile the binary output file
   * @exception IOException
   */
  public void write(File outputFile) throws IOException
  {
    // Thanks to Guenther for contributing the ExtSST implementation portion
    // of this method
    int charsLeft = 0;
    String curString = null;
    SSTRecord sst = new SSTRecord(totalOccurrences, stringList.size());
    ExtendedSSTRecord extsst = new ExtendedSSTRecord(stringList.size());
    int bucketSize = extsst.getNumberOfStringsPerBucket();

    Iterator i = stringList.iterator();
    int stringIndex = 0;
    while (i.hasNext() && charsLeft == 0)
    {
      curString = (String) i.next();
      // offset + header bytes
      int relativePosition = sst.getOffset() + 4;
      charsLeft = sst.add(curString);
      if ((stringIndex % bucketSize) == 0) {
        extsst.addString(outputFile.getPos(), relativePosition);
      }
      stringIndex++;
    }
    outputFile.write(sst);

    if (charsLeft != 0 || i.hasNext())
    {
      // Add the remainder of the string to the continue record
      SSTContinueRecord cont = createContinueRecord(curString,
                                                    charsLeft,
                                                    outputFile);

      // Carry on looping through the array until all the strings are done
      while (i.hasNext())
      {
        curString = (String) i.next();
        int relativePosition = cont.getOffset() + 4;
        charsLeft = cont.add(curString);
        if ((stringIndex % bucketSize) == 0) {
          extsst.addString(outputFile.getPos(), relativePosition);
        }
        stringIndex++;

        if (charsLeft != 0)
        {
          outputFile.write(cont);
          cont = createContinueRecord(curString, charsLeft, outputFile);
        }
      }

      outputFile.write(cont);
    }

    outputFile.write(extsst);
  }

  /**
   * Creates and returns a continue record using the left over bits and
   * pieces
   */
  private SSTContinueRecord createContinueRecord
    (String curString, int charsLeft, File outputFile) throws IOException
  {
    // Set up the remainder of the string in the continue record
    SSTContinueRecord cont = null;
    while (charsLeft != 0)
    {
      cont = new SSTContinueRecord();

      if (charsLeft == curString.length()  || curString.length() == 0)
      {
        charsLeft = cont.setFirstString(curString, true);
      }
      else
      {
        charsLeft = cont.setFirstString
          (curString.substring(curString.length() - charsLeft), false);
      }

      if (charsLeft != 0)
      {
        outputFile.write(cont);
        cont = new SSTContinueRecord();
      }
    }

    return cont;
  }
}
