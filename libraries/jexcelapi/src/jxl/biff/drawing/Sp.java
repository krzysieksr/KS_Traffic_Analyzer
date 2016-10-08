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
 * The Sp escher atom
 */
class Sp extends EscherAtom
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Sp.class);

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * The shape type
   */
  private int shapeType;

  /**
   * The shape id
   */
  private int shapeId;

  /**
   * The Sp persistence flags
   */
  private int persistenceFlags;

  /**
   * Constructor
   *
   * @param erd the entity record data
   */
  public Sp(EscherRecordData erd)
  {
    super(erd);
    shapeType = getInstance();
    byte[] bytes = getBytes();
    shapeId = IntegerHelper.getInt(bytes[0], bytes[1], bytes[2], bytes[3]);
    persistenceFlags = IntegerHelper.getInt(bytes[4], bytes[5],
                                           bytes[6], bytes[7]);
  }

  /**
   * Constructor - used when writing
   *
   * @param st the shape type
   * @param sid the shape id
   * @param p persistence flags
   */
  public Sp(ShapeType st, int sid, int p)
  {
    super(EscherRecordType.SP);
    setVersion(2);
    shapeType = st.getValue();
    shapeId = sid;
    persistenceFlags = p;
    setInstance(shapeType);
  }

  /**
   * Accessor for the shape id
   *
   * @return the shape id
   */
  int getShapeId()
  {
    return shapeId;
  }

  /**
   * Accessor for the shape type
   *
   * @return the shape type
   */
  int getShapeType()
  {
    return shapeType;
  }

  /**
   * Gets the data
   *
   * @return the binary data
   */
  byte[] getData()
  {
    data = new byte[8];
    IntegerHelper.getFourBytes(shapeId, data, 0);
    IntegerHelper.getFourBytes(persistenceFlags, data, 4);
    return setHeaderData(data);
  }
}
