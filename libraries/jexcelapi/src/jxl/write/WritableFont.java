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

package jxl.write;

import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.ScriptStyle;
import jxl.format.UnderlineStyle;
import jxl.write.biff.WritableFontRecord;

/**
 * A class which is instantiated when the user application wishes to specify
 * the font for a particular cell
 */
public class WritableFont extends WritableFontRecord
{
  /**
   * Static inner class used for classifying the font names
   */
  public static class FontName
  {
    /**
     * The name
     */
    String name;

    /**
     * Constructor
     *
     * @param s the font name
     */
    FontName(String s)
    {
      name = s;
    }
  }

  /**
   * Static inner class used for the boldness of the fonts
   */
  /*private*/ static class BoldStyle
  {
    /**
     * The value
     */
    public int value;

    /**
     * Constructor
     *
     * @param val the value
     */
    BoldStyle(int val)
    {
      value = val;
    }
  }

  /**
   * Objects created with this font name will be rendered within Excel as ARIAL
   * fonts
   */
  public static final FontName ARIAL = new FontName("Arial");
  /**
   * Objects created with this font name will be rendered within Excel as TIMES
   * fonts
   */
  public static final FontName TIMES = new FontName("Times New Roman");
  /**
   * Objects created with this font name will be rendered within Excel as
   * COURIER fonts
   */
  public static final FontName COURIER = new FontName("Courier New");
  /**
   * Objects created with this font name will be rendered within Excel as
   * TAHOMA fonts
   */
  public static final FontName TAHOMA = new FontName("Tahoma");

  // The bold styles

  /**
   * Indicates that this font should not be presented as bold
   */
  public static final BoldStyle NO_BOLD  = new BoldStyle(0x190);
  /**
   * Indicates that this font should be presented in a BOLD style
   */
  public static final BoldStyle BOLD     = new BoldStyle(0x2bc);

  /**
   * The default point size for all Fonts
   */
  public static final int DEFAULT_POINT_SIZE = 10;

  /**
   * Creates a default font, vanilla font of the specified face and with
   * default point size.
   *
   * @param fn the font name
   */
  public WritableFont(FontName fn)
  {
    this(fn,
         DEFAULT_POINT_SIZE,
         NO_BOLD,
         false,
         UnderlineStyle.NO_UNDERLINE,
         Colour.BLACK,
         ScriptStyle.NORMAL_SCRIPT);
  }

  /**
   * Publicly available copy constructor
   *
   * @param f the font to copy
   */
  public WritableFont(Font f)
  {
    super(f);
  }

  /**
   * Constructs of font of the specified face and of size given by the
   * specified point size
   *
   * @param ps the point size
   * @param fn the font name
   */
  public WritableFont(FontName fn, int ps)
  {
    this(fn, ps, NO_BOLD, false,
         UnderlineStyle.NO_UNDERLINE,
         Colour.BLACK,
         ScriptStyle.NORMAL_SCRIPT);
  }

  /**
   * Creates a font of the specified face, point size and bold style
   *
   * @param ps the point size
   * @param bs the bold style
   * @param fn the font name
   */
  public WritableFont(FontName fn, int ps, BoldStyle bs)
  {
    this(fn, ps, bs, false,
         UnderlineStyle.NO_UNDERLINE,
         Colour.BLACK,
         ScriptStyle.NORMAL_SCRIPT);
  }

  /**
   * Creates a font of the specified face, point size, bold weight and
   * italicised option.
   *
   * @param ps the point size
   * @param bs the bold style
   * @param italic italic flag
   * @param fn the font name
   */
  public WritableFont(FontName fn, int ps, BoldStyle bs, boolean italic)
  {
    this(fn, ps, bs, italic,
         UnderlineStyle.NO_UNDERLINE,
         Colour.BLACK,
         ScriptStyle.NORMAL_SCRIPT);
  }

  /**
   * Creates a font of the specified face, point size, bold weight,
   * italicisation and underline style
   *
   * @param ps the point size
   * @param bs the bold style
   * @param us the underline style
   * @param fn the font name
   * @param it italic flag
   */
  public WritableFont(FontName fn,
                      int ps,
                      BoldStyle bs,
                      boolean it,
                      UnderlineStyle us)
  {
    this(fn, ps, bs, it, us, Colour.BLACK, ScriptStyle.NORMAL_SCRIPT);
  }


  /**
   * Creates a font of the specified face, point size, bold style,
   * italicisation, underline style and colour
   *
   * @param ps the point size
   * @param bs the bold style
   * @param us the underline style
   * @param fn the font name
   * @param it italic flag
   * @param c the colour
   */
  public WritableFont(FontName fn,
                      int ps,
                      BoldStyle bs,
                      boolean it,
                      UnderlineStyle us,
                      Colour c)
  {
    this(fn, ps, bs, it, us, c, ScriptStyle.NORMAL_SCRIPT);
  }


  /**
   * Creates a font of the specified face, point size, bold style,
   * italicisation, underline style, colour, and script
   * style (superscript/subscript)
   *
   * @param ps the point size
   * @param bs the bold style
   * @param us the underline style
   * @param fn the font name
   * @param it the italic flag
   * @param c the colour
   * @param ss the script style
   */
  public WritableFont(FontName fn,
                      int ps,
                      BoldStyle bs,
                      boolean it,
                      UnderlineStyle us,
                      Colour c,
                      ScriptStyle ss)
  {
    super(fn.name, ps, bs.value, it,
          us.getValue(),
          c.getValue(), ss.getValue());
  }

  /**
   * Sets the point size for this font, if the font hasn't been initialized
   *
   * @param pointSize the point size
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setPointSize(int pointSize) throws WriteException
  {
    super.setPointSize(pointSize);
  }

  /**
   * Sets the bold style for this font, if the font hasn't been initialized
   *
   * @param boldStyle the bold style
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setBoldStyle(BoldStyle boldStyle) throws WriteException
  {
    super.setBoldStyle(boldStyle.value);
  }

  /**
   * Sets the italic indicator for this font, if the font hasn't been
   * initialized
   *
   * @param italic the italic flag
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setItalic(boolean italic) throws WriteException
  {
    super.setItalic(italic);
  }

  /**
   * Sets the underline style for this font, if the font hasn't been
   * initialized
   *
   * @param us the underline style
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setUnderlineStyle(UnderlineStyle us) throws WriteException
  {
    super.setUnderlineStyle(us.getValue());
  }

  /**
   * Sets the colour for this font, if the font hasn't been
   * initialized
   *
   * @param colour the colour
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setColour(Colour colour) throws WriteException
  {
    super.setColour(colour.getValue());
  }

  /**
   * Sets the script style (eg. superscript, subscript) for this font,
   * if the font hasn't been initialized
   *
   * @param scriptStyle the colour
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setScriptStyle(ScriptStyle scriptStyle) throws WriteException
  {
    super.setScriptStyle(scriptStyle.getValue());
  }

  /**
   * Accessor for the strike-out flag
   *
   * @return the strike-out flag
   */
  public boolean isStruckout()
  {
    return super.isStruckout();
  }

  /**
   * Sets Accessor for the strike-out flag
   *
   * @param struckout TRUE if this is a struckout font
   * @return the strike-out flag
   * @exception WriteException, if this font is already in use elsewhere
   */
  public void setStruckout(boolean struckout) throws WriteException
  {
    super.setStruckout(struckout);
  }

  /**
   * Factory method which creates the specified font name.  This method
   * should be used with care, since the string used to create the font
   * name must be recognized by Excel's internal processing
   *
   * @param fontName the name of the Excel font
   * @return the font name
   */
  public static FontName createFont(String fontName)
  {
    return new FontName(fontName);
  }
}

