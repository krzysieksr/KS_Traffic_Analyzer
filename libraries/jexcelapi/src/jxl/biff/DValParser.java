/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

import jxl.common.Logger;

import jxl.biff.IntegerHelper;

/**
 * Class which parses the binary data associated with Data Validity (DVal)
 * setting
 */
public class DValParser
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(DValParser.class);

  // The option masks
  private static int PROMPT_BOX_VISIBLE_MASK = 0x1;
  private static int PROMPT_BOX_AT_CELL_MASK = 0x2;
  private static int VALIDITY_DATA_CACHED_MASK = 0x4;

  /**
   * Prompt box visible
   */
  private boolean promptBoxVisible;

  /**
   * Empty cells allowed
   */
  private boolean promptBoxAtCell;

  /**
   * Cell validity data cached in following DV records
   */
  private boolean validityDataCached;

  /**
   * The number of following DV records
   */
  private int numDVRecords;

  /**
   * The object id of the associated down arrow
   */
  private int objectId;

  /**
   * Constructor
   */
  public DValParser(byte[] data)
  {
    int options = IntegerHelper.getInt(data[0], data[1]);

    promptBoxVisible = (options & PROMPT_BOX_VISIBLE_MASK) != 0;
    promptBoxAtCell = (options & PROMPT_BOX_AT_CELL_MASK) != 0;
    validityDataCached = (options & VALIDITY_DATA_CACHED_MASK) != 0;
    
    objectId = IntegerHelper.getInt(data[10], data[11], data[12], data[13]);
    numDVRecords = IntegerHelper.getInt(data[14], data[15], 
                                        data[16], data[17]);    
  }

  /**
   * Constructor
   */
  public DValParser(int objid, int num)
  {
    objectId = objid;
    numDVRecords = num;
    validityDataCached = true;
  }

  /**
   * Gets the data
   */
  public byte[] getData()
  {
    byte[] data = new byte[18];

    int options = 0;
    
    if (promptBoxVisible)
    {
      options |= PROMPT_BOX_VISIBLE_MASK;
    }

    if (promptBoxAtCell)
    {
      options |= PROMPT_BOX_AT_CELL_MASK;
    }

    if (validityDataCached)
    {
      options |= VALIDITY_DATA_CACHED_MASK;
    }

    IntegerHelper.getTwoBytes(options, data, 0);

    IntegerHelper.getFourBytes(objectId, data, 10);

    IntegerHelper.getFourBytes(numDVRecords, data, 14);

    return data;
  }

  /**
   * Called when a remove row or column results in one of DV records being 
   * removed
   */
  public void dvRemoved()
  {
    numDVRecords--;
  }

  /**
   * Accessor for the number of DV records
   *
   * @return the number of DV records for this list
   */
  public int getNumberOfDVRecords()
  {
    return numDVRecords;
  }

  /**
   * Accessor for the object id
   *
   * @return the object id
   */
  public int getObjectId()
  {
    return objectId;
  }

  /**
   * Called when adding a DV record on a copied DVal
   */
  public void dvAdded()
  {
    numDVRecords++;
  }
}
