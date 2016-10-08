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
 * The client anchor record
 */
class ClientAnchor extends EscherAtom
{
  /**
   * The logger
   */
  private static final Logger logger = Logger.getLogger(ClientAnchor.class);

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * The properties
   */
  private int properties;

  /**
   * The x1 position
   */
  private double x1;

  /**
   * The y1 position
   */
  private double y1;

  /**
   * The x2 position
   */
  private double x2;

  /**
   * The y2 position
   */
  private double y2;

  /**
   * Constructor
   *
   * @param erd the escher record data
   */
  public ClientAnchor(EscherRecordData erd)
  {
    super(erd);
    byte[] bytes = getBytes();

    // The properties
    properties = IntegerHelper.getInt(bytes[0], bytes[1]);

    // The x1 cell
    int x1Cell = IntegerHelper.getInt(bytes[2], bytes[3]);
    int x1Fraction = IntegerHelper.getInt(bytes[4], bytes[5]);

    x1 = x1Cell + (double) x1Fraction / (double) 1024;

    // The y1 cell
    int y1Cell = IntegerHelper.getInt(bytes[6], bytes[7]);
    int y1Fraction = IntegerHelper.getInt(bytes[8], bytes[9]);

    y1 = y1Cell + (double) y1Fraction / (double) 256;

    // The x2 cell
    int x2Cell = IntegerHelper.getInt(bytes[10], bytes[11]);
    int x2Fraction = IntegerHelper.getInt(bytes[12], bytes[13]);

    x2 = x2Cell + (double) x2Fraction / (double) 1024;

    // The y1 cell
    int y2Cell = IntegerHelper.getInt(bytes[14], bytes[15]);
    int y2Fraction = IntegerHelper.getInt(bytes[16], bytes[17]);

    y2 = y2Cell + (double) y2Fraction / (double) 256;
  }

  /**
   * Constructor
   *
   * @param x1 the x1 position
   * @param y1 the y1 position
   * @param x2 the x2 position
   * @param y2 the y2 position
   * @param props the anchor properties
   */
  public ClientAnchor(double x1, double y1, double x2, double y2, int props)
  {
    super(EscherRecordType.CLIENT_ANCHOR);
    this.x1 = x1;
    this.y1 = y1;
    this.x2 = x2;
    this.y2 = y2;
    properties = props;
  }

  /**
   * Gets the client anchor data
   *
   * @return the data
   */
  byte[] getData()
  {
    data = new byte[18];
    IntegerHelper.getTwoBytes(properties, data, 0);

    // The x1 cell
    IntegerHelper.getTwoBytes((int) x1, data, 2);

    // The x1 fraction into the cell 0-1024
    int x1fraction = (int) ((x1 - (int) x1) * 1024);
    IntegerHelper.getTwoBytes(x1fraction, data, 4);

    // The y1 cell
    IntegerHelper.getTwoBytes((int) y1, data, 6);

    // The y1 fraction into the cell 0-256
    int y1fraction = (int) ((y1 - (int) y1) * 256);
    IntegerHelper.getTwoBytes(y1fraction, data, 8);

    // The x2 cell
    IntegerHelper.getTwoBytes((int) x2, data, 10);

    // The x2 fraction into the cell 0-1024
    int x2fraction = (int) ((x2 - (int) x2) * 1024);
    IntegerHelper.getTwoBytes(x2fraction, data, 12);

    // The y2 cell
    IntegerHelper.getTwoBytes((int) y2, data, 14);

    // The y2 fraction into the cell 0-256
    int y2fraction = (int) ((y2 - (int) y2) * 256);
    IntegerHelper.getTwoBytes(y2fraction, data, 16);

    return setHeaderData(data);
  }

  /**
   * Accessor for the x1 position
   *
   * @return the x1 position
   */
  double getX1()
  {
    return x1;
  }

  /**
   * Accessor for the y1 position
   *
   * @return the y1 position
   */
  double getY1()
  {
    return y1;
  }

  /**
   * Accessor for the x2 position
   *
   * @return the x2 position
   */
  double getX2()
  {
    return x2;
  }

  /**
   * Accessor for the y2 position
   *
   * @return the y2 position
   */
  double getY2()
  {
    return y2;
  }

  /** 
   * Accessor for the anchor properties
   */
  int getProperties()
  {
    return properties;
  }
}
