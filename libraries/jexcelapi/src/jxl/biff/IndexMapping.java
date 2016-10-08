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

package jxl.biff;

import jxl.common.Logger;

/**
 * This class is a wrapper for a list of mappings between indices.
 * It is used when removing duplicate records and specifies the new
 * index for cells which have the duplicate format
 */
public final class IndexMapping
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(IndexMapping.class);

  /**
   * The array of new indexes for an old one
   */
  private int[] newIndices;

  /**
   * Constructor
   *
   * @param size the number of index numbers to be mapped
   */
  public IndexMapping(int size)
  {
    newIndices = new int[size];
  }

  /**
   * Sets a mapping
   * @param oldIndex the old index
   * @param newIndex the new index
   */
  public void setMapping(int oldIndex, int newIndex)
  {
    newIndices[oldIndex] = newIndex;
  }

  /**
   * Gets the new cell format index
   * @param oldIndex the existing index number
   * @return the new index number
   */
  public int getNewIndex(int oldIndex)
  {
    return newIndices[oldIndex];
  }
}
