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

/**
 * A BStoreContainer escher record
 */
class BStoreContainer extends EscherContainer
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(BStoreContainer.class);

  /**
   * The number of blips inside this container
   */
  private int numBlips;

  /**
   * Constructor used to instantiate this object when reading from an
   * escher stream
   *
   * @param erd the escher data
   */
  public BStoreContainer(EscherRecordData erd)
  {
    super(erd);
    numBlips = getInstance();
  }

  /**
   * Constructor used when writing out an escher record
   */
  public BStoreContainer()
  {
    super(EscherRecordType.BSTORE_CONTAINER);
  }

  /**
   * Sets the number of drawings in this container
   *
   * @param count the number of blips
   */
  void setNumBlips(int count)
  {
    numBlips = count;
    setInstance(numBlips);
  }

  /**
   * Accessor for the number of blips
   *
   * @return the number of blips
   */
  public int getNumBlips()
  {
    return numBlips;
  }

  /**
   * Accessor for the drawing
   *
   * @param i the index number of the drawing to return
   * @return the drawing
   */
  public BlipStoreEntry getDrawing(int i)
  {
    EscherRecord[] children = getChildren();
    BlipStoreEntry bse = (BlipStoreEntry) children[i];
    return bse;
  }
}
