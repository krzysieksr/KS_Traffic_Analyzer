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

import java.util.ArrayList;

import jxl.common.Logger;

import jxl.biff.IntegerHelper;

/**
 * Dgg record
 */
class Dgg extends EscherAtom
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Dgg.class);

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * The number of clusters
   */
  private int numClusters;

  /**
   * The maximum shape id
   */
  private int maxShapeId;

  /**
   * The number of shapes saved
   */
  private int shapesSaved;

  /**
   * The number of drawings saved
   */
  private int drawingsSaved;

  /**
   * The clusters
   */
  private ArrayList clusters;

  /**
   * The cluster structure
   */
  static final class Cluster
  {
    /**
     * The drawing group id
     */
    int drawingGroupId;

    /**
     * The something or other
     */
    int shapeIdsUsed;

    /**
     * Constructor
     *
     * @param dgId the drawing group id
     * @param sids the sids
     */
    Cluster(int dgId, int sids)
    {
      drawingGroupId = dgId;
      shapeIdsUsed = sids;
    }
  }

  /**
   * Constructor
   *
   * @param erd the read in data
   */
  public Dgg(EscherRecordData erd)
  {
    super(erd);
    clusters = new ArrayList();
    byte[] bytes = getBytes();
    maxShapeId = IntegerHelper.getInt
      (bytes[0], bytes[1], bytes[2], bytes[3]);
    numClusters = IntegerHelper.getInt
      (bytes[4], bytes[5], bytes[6], bytes[7]);
    shapesSaved = IntegerHelper.getInt
      (bytes[8], bytes[9], bytes[10], bytes[11]);
    drawingsSaved = IntegerHelper.getInt
      (bytes[12], bytes[13], bytes[14], bytes[15]);

    int pos = 16;
    for (int i = 0; i < numClusters; i++)
    {
      int dgId = IntegerHelper.getInt(bytes[pos], bytes[pos + 1]);
      int sids = IntegerHelper.getInt(bytes[pos + 2], bytes[pos + 3]);
      Cluster c = new Cluster(dgId, sids);
      clusters.add(c);
      pos += 4;
    }
  }

  /**
   * Constructor
   *
   * @param numShapes the number of shapes
   * @param numDrawings the number of drawings
   */
  public Dgg(int numShapes, int numDrawings)
  {
    super(EscherRecordType.DGG);
    shapesSaved = numShapes;
    drawingsSaved = numDrawings;
    clusters = new ArrayList();
  }

  /**
   * Adds a cluster to this record
   *
   * @param dgid the id
   * @param sids the sid
   */
  void addCluster(int dgid, int sids)
  {
    Cluster c = new Cluster(dgid, sids);
    clusters.add(c);
  }

  /**
   * Gets the data for writing out
   *
   * @return the binary data
   */
  byte[] getData()
  {
    numClusters = clusters.size();
    data = new byte[16 + numClusters * 4];

    // The max shape id
    IntegerHelper.getFourBytes(1024 + shapesSaved, data, 0);

    // The number of clusters
    IntegerHelper.getFourBytes(numClusters, data, 4);

    // The number of shapes saved
    IntegerHelper.getFourBytes(shapesSaved, data, 8);

    // The number of drawings saved
    //    IntegerHelper.getFourBytes(drawingsSaved, data, 12);
    IntegerHelper.getFourBytes(1, data, 12);

    int pos = 16;
    for (int i = 0; i < numClusters; i++)
    {
      Cluster c = (Cluster) clusters.get(i);
      IntegerHelper.getTwoBytes(c.drawingGroupId, data, pos);
      IntegerHelper.getTwoBytes(c.shapeIdsUsed, data, pos + 2);
      pos += 4;
    }

    return setHeaderData(data);
  }

  /**
   * Accessor for the number of shapes saved
   *
   * @return the number of shapes saved
   */
  int getShapesSaved()
  {
    return shapesSaved;
  }

  /**
   * Accessor for the number of drawings saved
   *
   * @return the number of drawings saved
   */
  int getDrawingsSaved()
  {
    return drawingsSaved;
  }

  /**
   * Accessor for a particular cluster
   *
   * @param i the cluster number
   * @return the cluster
   */
  Cluster getCluster(int i)
  {
    return (Cluster) clusters.get(i);
  }
}
