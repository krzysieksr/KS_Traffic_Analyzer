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

package jxl.biff.drawing;

import java.util.ArrayList;

import jxl.common.Assert;
import jxl.common.Logger;

/**
 * Class used to concatenate all the data for the various drawing objects
 * into one continuous stream
 */
public class DrawingData implements EscherStream
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(DrawingData.class);

  /**
   * The drawing data
   */
  private byte[] drawingData;

  /**
   * The number of drawings
   */
  private int numDrawings;

  /**
   * Initialized flag
   */
  private boolean initialized;

  /**
   * The spgr container. The contains the SpContainer for each drawing
   */
  private EscherRecord[] spContainers;

  /**
   * Constructor
   */
  public DrawingData()
  {
    numDrawings = 0;
    drawingData = null;
    initialized = false;
  }

  /**
   * Initialization
   */
  private void initialize()
  {
    EscherRecordData er = new EscherRecordData(this, 0);
    Assert.verify(er.isContainer());

    EscherContainer dgContainer  = new EscherContainer(er);
    EscherRecord[] children = dgContainer.getChildren();

    children = dgContainer.getChildren();
    // Dg dg = (Dg) children[0];

    EscherContainer spgrContainer = null;

    for (int i = 0; i < children.length && spgrContainer == null; i++)
    {
      EscherRecord child = children[i];
      if (child.getType() == EscherRecordType.SPGR_CONTAINER)
      {
        spgrContainer = (EscherContainer) child;
      }
    }
    Assert.verify(spgrContainer != null);

    EscherRecord[] spgrChildren = spgrContainer.getChildren();

    // See if any of the spgrChildren are SpgrContainer
    boolean nestedContainers = false;
    for (int i = 0; i < spgrChildren.length && !nestedContainers; i++)
    {
      if (spgrChildren[i].getType() == EscherRecordType.SPGR_CONTAINER)
      {
        nestedContainers = true;
      }
    }

    // If there are no nested containers, simply set the spContainer list
    // to be the list of children
    if (!nestedContainers)
    {
      spContainers = spgrChildren;
    }
    else
    {
      // Go through the hierarchy and dig out all the Sp containers
      ArrayList sps = new ArrayList();
      getSpContainers(spgrContainer, sps);
      spContainers = new EscherRecord[sps.size()];
      spContainers = (EscherRecord[]) sps.toArray(spContainers);
    }

    initialized = true;
  }

  /**
   * Gets the sp container from the internal data
   *
   * @param spgrContainer the spgr container
   * @param sps the list of sp records
   */
  private void getSpContainers(EscherContainer spgrContainer, ArrayList sps)
  {
    EscherRecord[] spgrChildren  = spgrContainer.getChildren();
    for (int i = 0; i < spgrChildren.length; i++)
    {
      if (spgrChildren[i].getType() == EscherRecordType.SP_CONTAINER)
      {
        sps.add(spgrChildren[i]);
      }
      else if (spgrChildren[i].getType() == EscherRecordType.SPGR_CONTAINER)
      {
        getSpContainers((EscherContainer) spgrChildren[i], sps);
      }
      else
      {
        logger.warn("Spgr Containers contains a record other than Sp/Spgr " +
                    "containers");
      }
    }
  }

  /**
   * Adds the byte stream to the drawing data
   *
   * @param data the data to add
   */
  public void addData(byte[] data)
  {
    addRawData(data);
    numDrawings++;
  }

  /**
   * Adds the data to the array without incrementing the drawing number.
   * This is used by comments, which for some bizarre and inexplicable
   * reason split out the data
   *
   * @param data the data to add
   */
  public void addRawData(byte[] data)
  {
    if (drawingData == null)
    {
      drawingData = data;
      return;
    }

    // Resize the array
    byte[] newArray = new byte[drawingData.length + data.length];
    System.arraycopy(drawingData, 0, newArray, 0, drawingData.length);
    System.arraycopy(data, 0, newArray, drawingData.length, data.length);
    drawingData = newArray;

    // Dirty up this object
    initialized = false;
  }

  /**
   * Accessor for the number of drawings
   *
   * @return the current count of drawings
   */
  final int getNumDrawings()
  {
    return numDrawings;
  }

  /**
   * Gets the sp container for the specified drawing number
   *
   * @param drawingNum the drawing number for which to return the spContainer
   * @return the spcontainer
   */
  EscherContainer getSpContainer(int drawingNum)
  {
    if (!initialized)
    {
      initialize();
    }

    if ( (drawingNum + 1) >= spContainers.length)
    {
      throw new DrawingDataException();
    }

    EscherContainer spContainer =
      (EscherContainer) spContainers[drawingNum + 1];

    Assert.verify(spContainer != null);

    return spContainer;
  }

  /**
   * Gets the data which was read in for the drawings
   *
   * @return the drawing data
   */
  public byte[] getData()
  {
    return drawingData;
  }
}
