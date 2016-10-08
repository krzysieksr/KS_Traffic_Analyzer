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
import jxl.read.biff.Record;

/**
 * Record containing the list of data validation settings for a given sheet
 */
public class DataValidityListRecord extends WritableRecordData
{
  private static Logger logger = Logger.getLogger
    (DataValidityListRecord.class);

  /**
   * The number of settings records associated with this list
   */
  private int numSettings;

  /**
   * The object id of the associated down arrow
   */
  private int objectId;

  /**
   * The dval parser
   */
  private DValParser dvalParser;

  /**
   * The data
   */
  private byte[] data;

  /**
   * Constructor
   */
  public DataValidityListRecord(Record t)
  {
    super(t);

    data = getRecord().getData();
    objectId = IntegerHelper.getInt(data[10], data[11], data[12], data[13]);
    numSettings = IntegerHelper.getInt(data[14], data[15], data[16], data[17]);
  }

  /**
   * Constructor called when generating a data validity list from the API
   */
  public DataValidityListRecord(DValParser dval)
  {
    super(Type.DVAL);

    dvalParser = dval;
  }

  /**
   * Copy constructor
   *
   * @param dvlr the record copied from a read only sheet
   */
  DataValidityListRecord(DataValidityListRecord dvlr)
  {
    super(Type.DVAL);

    data = dvlr.getData();
  }

  /**
   * Accessor for the number of settings records associated with this list
   */
  int getNumberOfSettings()
  {
    return numSettings;
  }

  /**
   * Retrieves the data for output to binary file
   * 
   * @return the data to be written
   */
  public byte[] getData()
  {
    if (dvalParser == null)
    {
      return data;
    }

    return dvalParser.getData();
  }

  /**
   * Called when a remove row or column results in one of DV records being 
   * removed
   */
  void dvRemoved()
  {
    if (dvalParser == null)
    {
      dvalParser = new DValParser(data);
    }

    dvalParser.dvRemoved();
  }

  /**
   * Called when a writable DV record is added to a copied validity list
   */
  void dvAdded()
  {
    if (dvalParser == null)
    {
      dvalParser = new DValParser(data);
    }

    dvalParser.dvAdded();
  }

  /**
   * Accessor for the number of DV records
   *
   * @return the number of DV records for this list
   */
  public boolean hasDVRecords()
  {
    if (dvalParser == null)
    {
      return true;
    }

    return dvalParser.getNumberOfDVRecords() > 0;
  }

  /**
   * Accessor for the object id
   *
   * @return the object id
   */
  public int getObjectId()
  {
    if (dvalParser == null)
    {
      return objectId;
    }
    
    return dvalParser.getObjectId();
  }
}
