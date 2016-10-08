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

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Indicates an extension to the Shared String Table.  Currently this
 * contains blank records
 *
 * Thanks to Guenther for contributing a proper implementation of the EXTSST
 * record, replacing my previous dummy version
 */
class ExtendedSSTRecord extends WritableRecordData
{
  private static final int infoRecordSize = 8;
  private int numberOfStrings;
  private int[] absoluteStreamPositions;
  private int[] relativeStreamPositions;
  private int currentStringIndex = 0;
  
  /**
   * Constructor
   *
   * @param numstrings the number of strings per bucket
   * @param streampos the absolute stream position of the beginning of
   * the SST record
   */
  public ExtendedSSTRecord(int newNumberOfStrings)
  {
    super(Type.EXTSST);
    numberOfStrings = newNumberOfStrings;
    int numberOfBuckets = getNumberOfBuckets();
    absoluteStreamPositions = new int[numberOfBuckets];
    relativeStreamPositions = new int[numberOfBuckets];
    currentStringIndex = 0;
  }

  public int getNumberOfBuckets() 
  {
    int numberOfStringsPerBucket = getNumberOfStringsPerBucket();
    return numberOfStringsPerBucket != 0 ?
      (numberOfStrings + numberOfStringsPerBucket - 1) / 
      numberOfStringsPerBucket : 0;
  }

  public int getNumberOfStringsPerBucket() 
  {
    // XXX
    // should come up with a more clever calculation
    // bucket limit should not be bigger than 1024, otherwise we end
    // up with too many buckets and would have to write continue records
    // for the EXTSST record which we want to avoid for now.
    final int bucketLimit = 128;
    return (numberOfStrings + bucketLimit - 1) / bucketLimit;
  }
  
  public void addString(int absoluteStreamPosition, 
                        int relativeStreamPosition) 
  {
    absoluteStreamPositions[currentStringIndex] = 
      absoluteStreamPosition + relativeStreamPosition;
    relativeStreamPositions[currentStringIndex] = relativeStreamPosition;
    currentStringIndex++;
  }

  /**
   * Gets the binary data to be written out
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    int numberOfBuckets = getNumberOfBuckets();
    byte[] data = new byte[2 + (8 * numberOfBuckets)];
    // number of strings per bucket
    IntegerHelper.getTwoBytes(getNumberOfStringsPerBucket(), data, 0);
    
    for (int i = 0; i < numberOfBuckets; i++) 
    {
      // absolute stream position
      IntegerHelper.getFourBytes(absoluteStreamPositions[i], 
                                 data, 
                                 2 + (i * infoRecordSize));
      // relative offset
      IntegerHelper.getTwoBytes(relativeStreamPositions[i], 
                                data, 
                                6 + (i * infoRecordSize));
      // reserved
      // IntegerHelper.getTwoBytes(0x0, data, 8 + (i * infoRecordSize));
    }

    return data;
  }
}

