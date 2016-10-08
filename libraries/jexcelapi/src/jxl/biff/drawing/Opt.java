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

import jxl.biff.IntegerHelper;
import jxl.biff.StringHelper;

/**
 * An options record in the escher stream
 */
class Opt extends EscherAtom
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Opt.class);

  /**
   * The binary data
   */
  private byte[] data;

  /**
   * The number of properties
   */
  private int numProperties;

  /**
   * The list of properties
   */
  private ArrayList properties;

  /**
   * Properties enumeration inner class
   */
  final static class Property
  {
    int id;
    boolean blipId;
    boolean complex;
    int value;
    String stringValue;

    /**
     * Constructor
     *
     * @param i the property id
     * @param bl the blip id
     * @param co complex flag
     * @param v the value
     */
    public Property(int i, boolean bl, boolean co, int v)
    {
      id = i;
      blipId = bl;
      complex = co;
      value = v;
    }

    /**
     * Constructor
     *
     * @param i the property id
     * @param bl the blip id
     * @param co complex flag
     * @param v the value
     * @param s the property string
     */
    public Property(int i, boolean bl, boolean co, int v, String s)
    {
      id = i;
      blipId = bl;
      complex = co;
      value = v;
      stringValue = s;
    }
  }

  /**
   * Constructor
   *
   * @param erd the escher record data
   */
  public Opt(EscherRecordData erd)
  {
    super(erd);
    numProperties = getInstance();
    readProperties();
  }

  /**
   * Reads the properties
   */
  private void readProperties()
  {
    properties = new ArrayList();
    int pos = 0;
    byte[] bytes = getBytes();

    for (int i = 0; i < numProperties; i++)
    {
      int val = IntegerHelper.getInt(bytes[pos], bytes[pos + 1]);
      int id = val & 0x3fff;
      int value = IntegerHelper.getInt(bytes[pos + 2], bytes[pos + 3],
                                       bytes[pos + 4], bytes[pos + 5]);
      Property p = new Property(id,
                                (val & 0x4000) != 0,
                                (val & 0x8000) != 0,
                                value);
      pos += 6;
      properties.add(p);
    }

    for (Iterator i = properties.iterator(); i.hasNext();)
    {
      Property p = (Property) i.next();
      if (p.complex)
      {
        p.stringValue = StringHelper.getUnicodeString(bytes, p.value / 2,
                                                      pos);
        pos += p.value;
      }
    }
  }

  /**
   * Constructor
   */
  public Opt()
  {
    super(EscherRecordType.OPT);
    properties = new ArrayList();
    setVersion(3);
  }

  /**
   * Accessor for the binary data
   *
   * @return the binary data
   */
  byte[] getData()
  {
    numProperties = properties.size();
    setInstance(numProperties);

    data = new byte[numProperties * 6];
    int pos = 0;

    // Add in the root data
    for (Iterator i = properties.iterator(); i.hasNext();)
    {
      Property p = (Property) i.next();
      int val = p.id & 0x3fff;

      if (p.blipId)
      {
        val |= 0x4000;
      }

      if (p.complex)
      {
        val |= 0x8000;
      }

      IntegerHelper.getTwoBytes(val, data, pos);
      IntegerHelper.getFourBytes(p.value, data, pos + 2);
      pos += 6;
    }

    // Add in any complex data
    for (Iterator i = properties.iterator(); i.hasNext();)
    {
      Property p = (Property) i.next();

      if (p.complex && p.stringValue != null)
      {
        byte[] newData =
          new byte[data.length + p.stringValue.length() * 2];
        System.arraycopy(data, 0, newData, 0, data.length);
        StringHelper.getUnicodeBytes(p.stringValue, newData, data.length);
        data = newData;
      }
    }

    return setHeaderData(data);
  }

  /**
   * Adds a property into the options
   *
   * @param id the property id
   * @param blip the blip id
   * @param complex whether it's a complex property
   * @param val the value
   */
  void addProperty(int id, boolean blip, boolean complex, int val)
  {
    Property p = new Property(id, blip, complex, val);
    properties.add(p);
  }

  /**
   * Adds a property into the options
   *
   * @param id the property id
   * @param blip the blip id
   * @param complex whether it's a complex property
   * @param val the value
   * @param s the value string
   */
  void addProperty(int id, boolean blip, boolean complex, int val, String s)
  {
    Property p = new Property(id, blip, complex, val, s);
    properties.add(p);
  }

  /**
   * Accessor for the property
   *
   * @param id the property id
   * @return the property
   */
  Property getProperty(int id)
  {
    boolean found = false;
    Property p = null;
    for (Iterator i = properties.iterator(); i.hasNext() && !found;)
    {
      p = (Property) i.next();
      if (p.id == id)
      {
        found = true;
      }
    }
    return found ? p : null;
  }
}
