/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

package jxl.biff.drawing;

import jxl.common.Logger;

import jxl.biff.IntegerHelper;


/**
 * A single record from an Escher stream.  Basically this a container for
 * the header data for each Escher record
 */
final class EscherRecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(EscherRecordData.class);

  /**
   * The byte position of this record in the escher stream
   */
  private int pos;

  /**
   * The instance value
   */
  private int instance;

  /**
   * The version value
   */
  private int version;

  /**
   * The record id
   */
  private int recordId;

  /**
   * The length of the record, excluding the 8 byte header
   */
  private int length;

  /**
   * The length of the stream
   */
  private int streamLength;

  /**
   * Indicates whether this record is a container
   */
  private boolean container;

  /**
   * The type of this record
   */
  private EscherRecordType type;

  /**
   * A handle back to the drawing group, which contains the entire escher
   * stream byte data
   */
  private EscherStream escherStream;

  /**
   * Constructor
   *
   * @param dg the escher stream data
   * @param p the current position in the stream
   */
  public EscherRecordData(EscherStream dg, int p)
  {
    escherStream = dg;
    pos = p;
    byte[] data = escherStream.getData();

    streamLength = data.length;

    // First two bytes contain instance and version
    int value = IntegerHelper.getInt(data[pos], data[pos + 1]);

    // Instance value is the first 12 bits
    instance = (value & 0xfff0) >> 4;

    // Version is the last four bits
    version = value & 0xf;

    // Bytes 2 and 3 are the record id
    recordId = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);

    // Length is bytes 4,5,6 and 7
    length = IntegerHelper.getInt(data[pos + 4], data[pos + 5],
                                  data[pos + 6], data[pos + 7]);

    if (version == 0x0f)
    {
      container = true;
    }
    else
    {
      container = false;
    }
  }

  /**
   * Constructor
   *
   * @param t the type of the escher record
   */
  public EscherRecordData(EscherRecordType t)
  {
    type = t;
    recordId = type.getValue();
  }

  /**
   * Determines whether this record is a container
   *
   * @return TRUE if this is a container, FALSE otherwise
   */
  public boolean isContainer()
  {
    return container;
  }

  /**
   * Accessor for the length, excluding the 8 byte header
   *
   * @return the length excluding the 8 byte header
   */
  public int getLength()
  {
    return length;
  }

  /**
   * Accessor for the record id
   *
   * @return the record id
   */
  public int getRecordId()
  {
    return recordId;
  }

  /**
   * Accessor for the drawing group stream
   *
   * @return the drawing group stream
   */
  EscherStream getDrawingGroup()
  {
    return escherStream;
  }

  /**
   * Gets the position in the stream
   *
   * @return the position in the stream
   */
  int getPos()
  {
    return pos;
  }

  /**
   * Gets the escher type of this record
   *
   * @return  the escher type
   */
  EscherRecordType getType()
  {
    if (type == null)
    {
      type = EscherRecordType.getType(recordId);
    }

    return type;
  }

  /**
   * Gets the instance value
   *
   * @return the instance value
   */
  int getInstance()
  {
    return instance;
  }

  /**
   * Sets whether or not this is a container - called when writing
   * out an escher stream
   *
   * @param c TRUE if this is a container, FALSE otherwise
   */
  void setContainer(boolean c)
  {
    container = c;
  }

  /**
   * Called from the subclass when writing to set the instance value
   *
   * @param inst the instance
   */
  void setInstance(int inst)
  {
    instance = inst;
  }

  /**
   * Called when writing to set the length of this record
   *
   * @param l the length
   */
  void setLength(int l)
  {
    length = l;
  }

  /**
   * Called when writing to set the version of this record
   *
   * @param v the version
   */
  void setVersion(int v)
  {
    version = v;
  }

  /**
   * Adds the 8 byte header data on the value data passed in, returning
   * the modified data
   *
   * @param d the value data
   * @return the value data with the header information
   */
  byte[] setHeaderData(byte[] d)
  {
    byte[] data = new byte[d.length + 8];
    System.arraycopy(d, 0, data, 8, d.length);

    if (container)
    {
      version = 0x0f;
    }

    // First two bytes contain instance and version
    int value = instance << 4;
    value |= version;
    IntegerHelper.getTwoBytes(value, data, 0);

    // Bytes 2 and 3 are the record id
    IntegerHelper.getTwoBytes(recordId, data, 2);

    // Length is bytes 4,5,6 and 7
    IntegerHelper.getFourBytes(d.length, data, 4);

    return data;
  }

  /**
   * Accessor for the header stream
   *
   * @return the escher stream
   */
  EscherStream getEscherStream()
  {
    return escherStream;
  }

  /**
   * Gets the data that was read in, excluding the header data
   *
   * @return the value data that was read in
   */
  byte[] getBytes()
  {
    byte[] d = new byte[length];
    System.arraycopy(escherStream.getData(), pos + 8, d, 0, length);
    return d;
  }

  /**
   * Accessor for the stream length
   *
   * @return the stream length
   */
  int getStreamLength()
  {
    return streamLength;
  }
}
