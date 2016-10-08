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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.ScriptStyle;
import jxl.format.UnderlineStyle;
import jxl.read.biff.Record;

/**
 * A record containing the necessary data for the font information
 */
public class FontRecord extends WritableRecordData implements Font
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(FontRecord.class);

  /**
   * The point height of this font
   */
  private int             pointHeight;
  /**
   * The index into the colour palette
   */
  private int             colourIndex;
  /**
   * The bold weight for this font (normal or bold)
   */
  private int             boldWeight;
  /**
   * The style of the script (italic or normal)
   */
  private int             scriptStyle;
  /**
   * The underline style for this font (none, single, double etc)
   */
  private int             underlineStyle;
  /**
   * The font family
   */
  private byte            fontFamily;
  /**
   * The character set
   */
  private byte            characterSet;

  /**
   * Indicates whether or not this font is italic
   */
  private boolean         italic;
  /**
   * Indicates whether or not this font is struck out
   */
  private boolean         struckout;
  /**
   * The name of this font
   */
  private String          name;
  /**
   * Flag to indicate whether the derived data (such as the font index) has
   * been initialized or not
   */
  private boolean         initialized;

  /**
   * The index of this font in the font list
   */
  private int fontIndex;

  /**
   * Dummy indicators for overloading the constructor
   */
  private static class Biff7 {};
  public static final Biff7 biff7 = new Biff7();

  /**
   * The conversion factor between microsoft internal units and point size
   */
  private static final int EXCEL_UNITS_PER_POINT = 20;

  /**
   * Constructor, used when creating a new font for writing out.
   *
   * @param bold the bold indicator
   * @param ps the point size
   * @param us the underline style
   * @param fn the name
   * @param it italicised indicator
   * @param ss the script style
   * @param ci the colour index
   */
  protected FontRecord(String fn, int ps, int bold, boolean it,
                       int us, int ci, int ss)
  {
    super(Type.FONT);
    boldWeight = bold;
    underlineStyle = us;
    name = fn;
    pointHeight = ps;
    italic = it;
    scriptStyle = ss;
    colourIndex = ci;
    initialized = false;
    struckout = false;
  }

  /**
   * Constructs this object from the raw data.  Used when reading in a
   * format record
   *
   * @param t the raw data
   * @param ws the workbook settings
   */
  public FontRecord(Record t, WorkbookSettings ws)
  {
    super(t);

    byte[] data = getRecord().getData();

    pointHeight = IntegerHelper.getInt(data[0], data[1]) /
      EXCEL_UNITS_PER_POINT;
    colourIndex = IntegerHelper.getInt(data[4], data[5]);
    boldWeight  = IntegerHelper.getInt(data[6], data[7]);
    scriptStyle = IntegerHelper.getInt(data[8], data[9]);
    underlineStyle = data[10];
    fontFamily   = data[11];
    characterSet = data[12];
    initialized  = false;

    if ((data[2] & 0x02) != 0)
    {
      italic = true;
    }

    if ((data[2] & 0x08) != 0)
    {
      struckout = true;
    }

    int numChars = data[14];
    if (data[15] == 0)
    {
      name = StringHelper.getString(data, numChars, 16, ws);
    }
    else if (data[15] == 1)
    {
      name = StringHelper.getUnicodeString(data, numChars, 16);
    }
    else
    {
      // Some font names don't have the unicode indicator
      name = StringHelper.getString(data, numChars, 15, ws);
    }
  }

  /**
   * Constructs this object from the raw data.  Used when reading in a
   * format record
   *
   * @param t the raw data
   * @param ws the workbook settings
   * @param dummy dummy overload
   */
  public FontRecord(Record t, WorkbookSettings ws, Biff7 dummy)
  {
    super(t);

    byte[] data = getRecord().getData();

    pointHeight = IntegerHelper.getInt(data[0], data[1]) /
      EXCEL_UNITS_PER_POINT;
    colourIndex = IntegerHelper.getInt(data[4], data[5]);
    boldWeight  = IntegerHelper.getInt(data[6], data[7]);
    scriptStyle = IntegerHelper.getInt(data[8], data[9]);
    underlineStyle = data[10];
    fontFamily = data[11];
    initialized = false;

    if ((data[2] & 0x02) != 0)
    {
      italic = true;
    }

    if ((data[2] & 0x08) != 0)
    {
      struckout = true;
    }

    int numChars = data[14];
    name = StringHelper.getString(data, numChars, 15, ws);
  }

  /**
   * Publicly available copy constructor
   *
   * @param  f the font to copy
   */
  protected FontRecord(Font f)
  {
    super(Type.FONT);
   
    Assert.verify(f != null);

    pointHeight = f.getPointSize();
    colourIndex = f.getColour().getValue();
    boldWeight = f.getBoldWeight();
    scriptStyle = f.getScriptStyle().getValue();
    underlineStyle = f.getUnderlineStyle().getValue();
    italic = f.isItalic();
    name = f.getName();
    struckout = f.isStruckout();
    initialized = false;
  }

  /**
   * Gets the byte data for writing out
   *
   * @return the raw data
   */
  public byte[] getData()
  {
    byte[] data = new byte[16 + name.length() * 2];

    // Excel expects font heights in 1/20ths of a point
    IntegerHelper.getTwoBytes(pointHeight * EXCEL_UNITS_PER_POINT, data, 0);

    // Set the font attributes to be zero for now
    if (italic)
    {
      data[2] |= 0x2;
    }

    if (struckout)
    {
      data[2] |= 0x08;
    }

    // Set the index to the colour palette
    IntegerHelper.getTwoBytes(colourIndex, data, 4);

    // Bold style
    IntegerHelper.getTwoBytes(boldWeight, data, 6);

    // Script style
    IntegerHelper.getTwoBytes(scriptStyle, data, 8);

    // Underline style
    data[10] = (byte) underlineStyle;

    // Set the font family to be 0
    data[11] = fontFamily;

    // Set the character set to be zero
    data[12] = characterSet;

    // Set the reserved bit to be zero
    data[13] = 0;

    // Set the length of the font name
    data[14] = (byte) name.length();

    data[15] = (byte) 1;

    // Copy in the string
    StringHelper.getUnicodeBytes(name, data, 16);

    return data;
  }

  /**
   * Accessor to see whether this object is initialized or not.
   *
   * @return TRUE if this font record has been initialized, FALSE otherwise
   */
  public final boolean isInitialized()
  {
    return initialized;
  }

  /**
   * Sets the font index of this record.  Called from the FormattingRecords
   *  object
   *
   * @param pos the position of this font in the workbooks font list
   */
  public final void initialize(int pos)
  {
    fontIndex = pos;
    initialized = true;
  }

  /**
   * Resets the initialize flag.  This is called by the constructor of
   * WritableWorkbookImpl to reset the statically declared fonts
   */
  public final void uninitialize()
  {
    initialized = false;
  }

  /**
   * Accessor for the font index
   *
   * @return the font index
   */
  public final int getFontIndex()
  {
    return fontIndex;
  }

  /**
   * Sets the point size for this font, if the font hasn't been initialized
   *
   * @param ps the point size
   */
  protected void setFontPointSize(int ps)
  {
    Assert.verify(!initialized);

    pointHeight = ps;
  }

  /**
   * Gets the point size for this font, if the font hasn't been initialized
   *
   * @return the point size
   */
  public int getPointSize()
  {
    return pointHeight;
  }

  /**
   * Sets the bold style for this font, if the font hasn't been initialized
   *
   * @param bs the bold style
   */
  protected void setFontBoldStyle(int bs)
  {
    Assert.verify(!initialized);

    boldWeight = bs;
  }

  /**
   * Gets the bold weight for this font
   *
   * @return the bold weight for this font
   */
  public int getBoldWeight()
  {
    return boldWeight;
  }

  /**
   * Sets the italic indicator for this font, if the font hasn't been
   * initialized
   *
   * @param i the italic flag
   */
  protected void setFontItalic(boolean i)
  {
    Assert.verify(!initialized);

    italic = i;
  }

  /**
   * Returns the italic flag
   *
   * @return TRUE if this font is italic, FALSE otherwise
   */
  public boolean isItalic()
  {
    return italic;
  }

  /**
   * Sets the underline style for this font, if the font hasn't been
   * initialized
   *
   * @param us the underline style
   */
  protected void setFontUnderlineStyle(int us)
  {
    Assert.verify(!initialized);

    underlineStyle = us;
  }

  /**
   * Gets the underline style for this font
   *
   * @return the underline style
   */
  public UnderlineStyle getUnderlineStyle()
  {
    return UnderlineStyle.getStyle(underlineStyle);
  }

  /**
   * Sets the colour for this font, if the font hasn't been
   * initialized
   *
   * @param c the colour
   */
  protected void setFontColour(int c)
  {
    Assert.verify(!initialized);

    colourIndex = c;
  }

  /**
   * Gets the colour for this font
   *
   * @return the colour
   */
  public Colour getColour()
  {
    return Colour.getInternalColour(colourIndex);
  }

  /**
   * Sets the script style (eg. superscript, subscript) for this font,
   * if the font hasn't been initialized
   *
   * @param ss the colour
   */
  protected void setFontScriptStyle(int ss)
  {
    Assert.verify(!initialized);

    scriptStyle = ss;
  }

  /**
   * Gets the script style
   *
   * @return the script style
   */
  public ScriptStyle getScriptStyle()
  {
    return ScriptStyle.getStyle(scriptStyle);
  }

  /**
   * Gets the name of this font
   *
   * @return the name of this font
   */
  public String getName()
  {
    return name;
  }

  /**
   * Standard hash code method
   * @return the hash code for this object
   */
  public int hashCode()
  {
    return name.hashCode();
  }

  /**
   * Standard equals method
   * @param o the object to compare
   * @return TRUE if the objects are equal, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof FontRecord))
    {
      return false;
    }

    FontRecord font = (FontRecord) o;

    if (pointHeight    == font.pointHeight &&
        colourIndex    == font.colourIndex &&
        boldWeight     == font.boldWeight &&
        scriptStyle    == font.scriptStyle &&
        underlineStyle == font.underlineStyle &&
        italic         == font.italic &&
        struckout      == font.struckout &&
        fontFamily     == font.fontFamily &&
        characterSet   == font.characterSet &&
        name.equals(font.name))
    {
      return true;
    }

    return false;
  }

  /**
   * Accessor for the strike out flag
   *
   * @return TRUE if this font is struck out, FALSE otherwise
   */
  public boolean isStruckout()
  {
    return struckout;
  }

  /**
   * Sets the struck out flag
   *
   * @param os TRUE if the font is struck out, false otherwise
   */
  protected void setFontStruckout(boolean os)
  {
    struckout = os;
  }
}


