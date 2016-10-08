/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan
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

import jxl.format.Colour;
import jxl.format.RGB;
import jxl.read.biff.Record;

/**
 * A record representing the RGB colour palette
 */
public class PaletteRecord extends WritableRecordData
{
  /**
   * The list of bespoke rgb colours used by this sheet
   */
  private RGB[] rgbColours = new RGB[numColours];

  /**
   * A dirty flag indicating that this palette has been tampered with
   * in some way
   */
  private boolean dirty;

  /**
   * Flag indicating that the palette was read in
   */
  private boolean read;

  /**
   * Initialized flag
   */
  private boolean initialized;

  /**
   * The number of colours in the palette
   */
  private static final int numColours = 56;

  /**
   * Constructor
   *
   * @param t the raw bytes
   */
  public PaletteRecord(Record t)
  {
    super(t);

    initialized = false;
    dirty = false;
    read = true;
  }

  /**
   * Default constructor - used when there is no palette specified
   */
  public PaletteRecord()
  {
    super(Type.PALETTE);

    initialized = true;
    dirty = false;
    read = false;

    // Initialize the array with all the default colours
    Colour[] colours = Colour.getAllColours();

    for (int i = 0; i < colours.length; i++)
    {
      Colour c = colours[i];
      setColourRGB(c,
                   c.getDefaultRGB().getRed(),
                   c.getDefaultRGB().getGreen(),
                   c.getDefaultRGB().getBlue());
    }
  }

  /**
   * Accessor for the binary data - used when copying
   *
   * @return the binary data
   */
  public byte[] getData()
  {
    // Palette was read in, but has not been changed
    if (read && !dirty)
    {
      return getRecord().getData();
    }

    byte[] data = new byte[numColours * 4 + 2];
    int pos = 0;

    // Set the number of records
    IntegerHelper.getTwoBytes(numColours, data, pos);

    // Set the rgb content
    for (int i = 0; i < numColours; i++)
    {
      pos = i * 4 + 2;
      data[pos]     = (byte) rgbColours[i].getRed();
      data[pos + 1] = (byte) rgbColours[i].getGreen();
      data[pos + 2] = (byte) rgbColours[i].getBlue();
    }

    return data;
  }

  /**
   * Initialize the record data
   */
  private void initialize()
  {
    byte[] data = getRecord().getData();

    int numrecords = IntegerHelper.getInt(data[0], data[1]);

    for (int i = 0; i < numrecords; i++)
    {
      int pos = i * 4 + 2;
      int red   = IntegerHelper.getInt(data[pos], (byte) 0);
      int green = IntegerHelper.getInt(data[pos + 1], (byte) 0);
      int blue  = IntegerHelper.getInt(data[pos + 2], (byte) 0);
      rgbColours[i] = new RGB(red, green, blue);
    }

    initialized = true;
  }

  /**
   * Accessor for the dirty flag, which indicates if this palette has been
   * modified
   *
   * @return TRUE if the palette has been modified, FALSE if it is the default
   */
  public boolean isDirty()
  {
    return dirty;
  }

  /**
   * Sets the RGB value for the specified colour for this workbook
   *
   * @param c the colour whose RGB value is to be overwritten
   * @param r the red portion to set (0-255)
   * @param g the green portion to set (0-255)
   * @param b the blue portion to set (0-255)
   */
  public void setColourRGB(Colour c, int r, int g, int b)
  {
    // Only colours on the standard palette with values 8-64 are acceptable
    int pos = c.getValue() - 8;
    if (pos < 0 || pos >= numColours)
    {
      return;
    }

    if (!initialized)
    {
      initialize();
    }

    // Force the colours into the range 0-255
    r = setValueRange(r, 0, 0xff);
    g = setValueRange(g, 0, 0xff);
    b = setValueRange(b, 0, 0xff);

    rgbColours[pos] = new RGB(r, g, b);

    // Indicate that the palette has been modified
    dirty = true;
  }

  /**
   * Gets the colour RGB from the palette
   *
   * @param c the colour
   * @return an RGB structure
   */
  public RGB getColourRGB(Colour c)
  {
    // Only colours on the standard palette with values 8-64 are acceptable
    int pos = c.getValue() - 8;
    if (pos < 0 || pos >= numColours)
    {
      return c.getDefaultRGB();
    }

    if (!initialized)
    {
      initialize();
    }

    return rgbColours[pos];
  }

  /**
   * Forces the value passed in to be between the range passed in
   *
   * @param val the value to constrain
   * @param min the minimum acceptable value
   * @param max the maximum acceptable value
   * @return the constrained value
   */
  private int setValueRange(int val, int min, int max)
  {
    val = Math.max(val, min);
    val = Math.min(val, max);
    return val;
  }
}
