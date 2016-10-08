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
import java.util.Iterator;

import jxl.common.Logger;

/**
 * An escher container.  This record may contain other escher containers or
 * atoms
 */
class EscherContainer extends EscherRecord
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(EscherContainer.class);

  /**
   * Initialized flag
   */
  private boolean initialized;


  /**
   * The children of this container
   */
  private ArrayList children;

  /**
   * Constructor
   *
   * @param erd the raw data
   */
  public EscherContainer(EscherRecordData erd)
  {
    super(erd);
    initialized = false;
    children = new ArrayList();
  }

  /**
   * Constructor used when writing out escher data
   *
   * @param type the type
   */
  protected EscherContainer(EscherRecordType type)
  {
    super(type);
    setContainer(true);
    children = new ArrayList();
  }

  /**
   * Accessor for the children of this container
   *
   * @return the children
   */
  public EscherRecord[] getChildren()
  {
    if (!initialized)
    {
      initialize();
    }

    Object[] ca = children.toArray(new EscherRecord[children.size()]);

    return (EscherRecord[]) ca;
  }

  /**
   * Adds a child to this container
   *
   * @param child the item to add
   */
  public void add(EscherRecord child)
  {
    children.add(child);
  }

  /**
   * Removes a child from this container
   *
   * @param child the item to remove
   */
  public void remove(EscherRecord child)
  {
    boolean result = children.remove(child);
  }

  /**
   * Initialization
   */
  private void initialize()
  {
    int curpos = getPos() + HEADER_LENGTH;
    int endpos = Math.min(getPos() + getLength(), getStreamLength());

    EscherRecord newRecord = null;

    while (curpos < endpos)
    {
      EscherRecordData erd = new EscherRecordData(getEscherStream(), curpos);

      EscherRecordType type = erd.getType();
      if (type == EscherRecordType.DGG)
      {
        newRecord = new Dgg(erd);
      }
      else if (type == EscherRecordType.DG)
      {
        newRecord = new Dg(erd);
      }
      else if (type == EscherRecordType.BSTORE_CONTAINER)
      {
        newRecord = new BStoreContainer(erd);
      }
      else if (type == EscherRecordType.SPGR_CONTAINER)
      {
        newRecord = new SpgrContainer(erd);
      }
      else if (type == EscherRecordType.SP_CONTAINER)
      {
        newRecord = new SpContainer(erd);
      }
      else if (type == EscherRecordType.SPGR)
      {
        newRecord = new Spgr(erd);
      }
      else if (type == EscherRecordType.SP)
      {
        newRecord = new Sp(erd);
      }
      else if (type == EscherRecordType.CLIENT_ANCHOR)
      {
        newRecord = new ClientAnchor(erd);
      }
      else if (type == EscherRecordType.CLIENT_DATA)
      {
        newRecord = new ClientData(erd);
      }
      else if (type == EscherRecordType.BSE)
      {
        newRecord = new BlipStoreEntry(erd);
      }
      else if (type == EscherRecordType.OPT)
      {
        newRecord = new Opt(erd);
      }
      else if (type == EscherRecordType.SPLIT_MENU_COLORS)
      {
        newRecord = new SplitMenuColors(erd);
      }
      else if (type == EscherRecordType.CLIENT_TEXT_BOX)
      {
        newRecord = new ClientTextBox(erd);
      }
      else
      {
        newRecord = new EscherAtom(erd);
      }

      children.add(newRecord);
      curpos += newRecord.getLength();
    }

    initialized = true;
  }

  /**
   * Gets the data for this container (and all of its children recursively
   *
   * @return the binary data
   */
  byte[] getData()
  {
    if (!initialized)
    {
      initialize();
    }

    byte[] data = new byte[0];
    for (Iterator i = children.iterator(); i.hasNext();)
    {
      EscherRecord er = (EscherRecord) i.next();
      byte[] childData = er.getData();

      if (childData != null)
      {
        byte[] newData = new byte[data.length + childData.length];
        System.arraycopy(data, 0, newData, 0, data.length);
        System.arraycopy(childData, 0, newData, data.length, childData.length);
        data = newData;
      }
    }

    return setHeaderData(data);
  }
}
