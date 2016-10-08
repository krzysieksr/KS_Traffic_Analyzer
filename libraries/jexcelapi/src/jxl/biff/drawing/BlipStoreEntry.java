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

import java.io.IOException;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.biff.IntegerHelper;

/**
 * The data for this blip store entry.  Typically this is the raw image data
 */
class BlipStoreEntry extends EscherAtom
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(BlipStoreEntry.class);

  /**
   * The type of the blip
   */
  private BlipType type;

  /**
   * The image data read in
   */
  private byte[] data;

  /**
   * The length of the image data
   */
  private int imageDataLength;

  /**
   * The reference count on this blip
   */
  private int referenceCount;

  /**
   * Flag to indicate that this entry was specified by the API, and not
   * read in
   */
  private boolean write;

  /**
   * The start of the image data within this blip entry
   */
  private static final int IMAGE_DATA_OFFSET = 61;

  /**
   * Constructor
   *
   * @param erd the escher record data
   */
  public BlipStoreEntry(EscherRecordData erd)
  {
    super(erd);
    type = BlipType.getType(getInstance());
    write = false;
    byte[] bytes = getBytes();
    referenceCount =  IntegerHelper.getInt(bytes[24], bytes[25],
                                           bytes[26], bytes[27]);
  }

  /**
   * Constructor
   *
   * @param d the drawing
   * @exception IOException
   */
  public BlipStoreEntry(Drawing d) throws IOException
  {
    super(EscherRecordType.BSE);
    type = BlipType.PNG;
    setVersion(2);
    setInstance(type.getValue());

    byte[] imageData = d.getImageBytes();
    imageDataLength = imageData.length;
    data = new byte[imageDataLength + IMAGE_DATA_OFFSET];
    System.arraycopy(imageData, 0, data, IMAGE_DATA_OFFSET, imageDataLength);
    referenceCount = d.getReferenceCount();
    write = true;
  }

  /**
   * Accessor for the blip type
   *
   * @return the blip type
   */
  public BlipType getBlipType()
  {
    return type;
  }

  /**
   * Gets the data for this blip so that it can be written out
   *
   * @return the data for the blip
   */
  public byte[] getData()
  {
    if (write)
    {
      // Drawing has been specified by API

      // Type on win32
      data[0] = (byte) type.getValue();

      // Type on MacOs
      data[1] = (byte) type.getValue();

      // The blip identifier
      //    IntegerHelper.getTwoBytes(0xfce1, data, 2);

      // Unused tags - 18 bytes
      //    System.arraycopy(stuff, 0, data, 2, stuff.length);

      // The size of the file
      IntegerHelper.getFourBytes(imageDataLength + 8 + 17, data, 20);

      // The reference count on the blip
      IntegerHelper.getFourBytes(referenceCount, data, 24);

      // Offset in the delay stream
      IntegerHelper.getFourBytes(0, data, 28);

      // Usage byte
      data[32] = (byte) 0;

      // Length of the blip name
      data[33] = (byte) 0;

      // Last two bytes unused
      data[34] = (byte) 0x7e;
      data[35] = (byte) 0x01;

      // The blip itself
      data[36] = (byte) 0;
      data[37] = (byte) 0x6e;

      // The blip identifier
      IntegerHelper.getTwoBytes(0xf01e, data, 38);

      // The length of the blip.  This is the length of the image file plus
      // 16 bytes
      IntegerHelper.getFourBytes(imageDataLength + 17, data, 40);

      // Unknown stuff
      //    System.arraycopy(stuff, 0, data, 44, stuff.length);
    }
    else
    {
      // drawing has been read in
      data = getBytes();
    }

    return setHeaderData(data);
  }

  /**
   * Reduces the reference count in this blip.  Called when a drawing is
   * removed
   */
  void dereference()
  {
    referenceCount--;
    Assert.verify(referenceCount >= 0);
  }

  /**
   * Accessor for the reference count on the blip
   *
   * @return the reference count on the blip
   */
  int getReferenceCount()
  {
    return referenceCount;
  }

  /**
   * Accessor for the image data.
   *
   * @return the image data
   */
  byte[] getImageData()
  {
    byte[] allData = getBytes();
    byte[] imageData = new byte[allData.length - IMAGE_DATA_OFFSET];
    System.arraycopy(allData, IMAGE_DATA_OFFSET,
                     imageData, 0, imageData.length);
    return imageData;
  }
}
