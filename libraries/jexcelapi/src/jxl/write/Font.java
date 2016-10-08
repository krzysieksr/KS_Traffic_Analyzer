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
import jxl.format.ScriptStyle;
import jxl.format.UnderlineStyle;

/**
 * A class which is instantiated when the user application wishes to specify
 * the font for a particular cell
 *
 * @deprecated Renamed to writable font
 */
public class Font extends WritableFont
{
  /**
   * Objects created with this font name will be rendered within Excel as ARIAL
   * fonts
   * @deprecated
   */
  public static final FontName ARIAL = WritableFont.ARIAL;
  /**
   * Objects created with this font name will be rendered within Excel as TIMES
   * fonts
   * @deprecated
   */
  public static final FontName TIMES = WritableFont.TIMES;

  // The bold styles

  /**
   * Indicates that this font should not be presented as bold
   * @deprecated
   */
  public static final BoldStyle NO_BOLD  = WritableFont.NO_BOLD;
  /**
   * Indicates that this font should be presented in a BOLD style
   * @deprecated
   */
  public static final BoldStyle BOLD     = WritableFont.BOLD;

  // The underline styles
  /**
   * @deprecated
   */
  public static final UnderlineStyle NO_UNDERLINE  =
    UnderlineStyle.NO_UNDERLINE;

  /**
   * @deprecated
   */
  public static final UnderlineStyle SINGLE        = UnderlineStyle.SINGLE;

  /**
   * @deprecated
   */
  public static final UnderlineStyle DOUBLE        = UnderlineStyle.DOUBLE;

  /**
   * @deprecated
   */
  public static final UnderlineStyle SINGLE_ACCOUNTING =
    UnderlineStyle.SINGLE_ACCOUNTING;

  /**
   * @deprecated
   */
  public static final UnderlineStyle DOUBLE_ACCOUNTING =
    UnderlineStyle.DOUBLE_ACCOUNTING;

  // The script styles
  public static final ScriptStyle NORMAL_SCRIPT = ScriptStyle.NORMAL_SCRIPT;
  public static final ScriptStyle SUPERSCRIPT   = ScriptStyle.SUPERSCRIPT;
  public static final ScriptStyle SUBSCRIPT     = ScriptStyle.SUBSCRIPT;

  /**
   * Creates a default font, vanilla font of the specified face and with
   * default point size.
   *
   * @param fn the font name
   * @deprecated Use jxl.write.WritableFont
   */
  public Font(FontName fn)
  {
    super(fn);
  }

  /**
   * Constructs of font of the specified face and of size given by the
   * specified point size
   *
   * @param ps the point size
   * @param fn the font name
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn, int ps)
  {
    super(fn, ps);
  }

  /**
   * Creates a font of the specified face, point size and bold style
   *
   * @param ps the point size
   * @param bs the bold style
   * @param fn the font name
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn, int ps, BoldStyle bs)
  {
    super(fn, ps, bs);
  }

  /**
   * Creates a font of the specified face, point size, bold weight and
   * italicised option.
   *
   * @param ps the point size
   * @param bs the bold style
   * @param italic italic flag
   * @param fn the font name
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn, int ps, BoldStyle bs, boolean italic)
  {
    super(fn, ps, bs, italic);
  }

  /**
   * Creates a font of the specified face, point size, bold weight,
   * italicisation and underline style
   *
   * @param ps the point size
   * @param bs the bold style
   * @param us underscore flag
   * @param fn font name
   * @param it italic flag
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn,
              int ps,
              BoldStyle bs,
              boolean it,
              UnderlineStyle us)
  {
    super(fn, ps, bs, it, us);
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
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn,
              int ps,
              BoldStyle bs,
              boolean it,
              UnderlineStyle us,
              Colour c)
  {
    super(fn, ps, bs, it, us, c);
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
   * @deprecated use jxl.write.WritableFont
   */
  public Font(FontName fn,
              int ps,
              BoldStyle bs,
              boolean it,
              UnderlineStyle us,
              Colour c,
              ScriptStyle ss)
  {
    super(fn, ps, bs, it, us, c, ss);
  }
}

