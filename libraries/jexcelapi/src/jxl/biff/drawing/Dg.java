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

import jxl.biff.IntegerHelper;

/**
 * The Drawing Group
 */
class Dg extends EscherAtom
{
  /**
   * The data
   */
  private byte[] data;

  /**
   * The id of this drawing
   */
  private int drawingId;

  /**
   * The number of shapes
   */
  private int shapeCount;

  /**
   * The seed for drawing ids
   */
  private int seed;

  /**
   * Constructor invoked when reading in an escher stream
   *
   * @param erd the escher record
   */
  public Dg(EscherRecordData erd)
  {
    super(erd);
    drawingId = getInstance();

    byte[] bytes = getBytes();
    shapeCount = IntegerHelper.getInt(bytes[0], bytes[1], bytes[2], bytes[3]);
    seed = IntegerHelper.getInt(bytes[4], bytes[5], bytes[6], bytes[7]);
  }

  /**
   * Constructor invoked when writing out an escher stream
   *
   * @param numDrawings the number of drawings
   */
  public Dg(int numDrawings)
  {
    super(EscherRecordType.DG);
    drawingId = 1;
    shapeCount = numDrawings + 1;
    seed = 1024 + shapeCount + 1;
    setInstance(drawingId);
  }

  /**
   * Gets the drawing id
   *
   * @return the drawing id
   */
  public int getDrawingId()
  {
    return drawingId;
  }

  /**
   * Gets the shape count
   *
   * @return the shape count
   */
  int getShapeCount()
  {
    return shapeCount;
  }

  /**
   * Used to generate the drawing data
   *
   * @return the data
   */
  byte[] getData()
  {
    data = new byte[8];
    IntegerHelper.getFourBytes(shapeCount, data, 0);
    IntegerHelper.getFourBytes(seed, data, 4);

    return setHeaderData(data);
  }
}
