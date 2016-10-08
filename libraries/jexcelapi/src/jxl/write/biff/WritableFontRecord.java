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

package jxl.write.biff;

import jxl.biff.FontRecord;
import jxl.format.Font;
import jxl.write.WriteException;

/**
 * A writable Font record.  This class intercepts any set accessor calls 
 * and throws and exception if the Font is already initialized
 */
public class WritableFontRecord extends FontRecord
{
  /**
   * Constructor, used when creating a new font for writing out.
   * 
   * @param bold the bold indicator
   * @param ps the point size
   * @param us the underline style
   * @param fn the name
   * @param it italicised indicator
   * @param c  the colour
   * @param ss the script style
   */
  protected WritableFontRecord(String fn, int ps, int bold, boolean it, 
                         int us, int ci, int ss)
  {
    super(fn, ps, bold, it, us, ci, ss);
  }

  /**
   * Publicly available copy constructor
   *
   * @param the font to copy
   */
  protected WritableFontRecord(Font f)
  {
    super(f);
  }


  /**
   * Sets the point size for this font, if the font hasn't been initialized
   * 
   * @param pointSize the point size
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setPointSize(int pointSize) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontPointSize(pointSize);
  }

  /**
   * Sets the bold style for this font, if the font hasn't been initialized
   * 
   * @param boldStyle the bold style
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setBoldStyle(int boldStyle) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontBoldStyle(boldStyle);
  }

  /**
   * Sets the italic indicator for this font, if the font hasn't been 
   * initialized
   * 
   * @param italic the italic flag
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setItalic(boolean italic) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontItalic(italic);
  }

  /**
   * Sets the underline style for this font, if the font hasn't been 
   * initialized
   * 
   * @param us the underline style
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setUnderlineStyle(int us) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontUnderlineStyle(us);
  }

  /**
   * Sets the colour for this font, if the font hasn't been 
   * initialized
   * 
   * @param colour the colour
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setColour(int colour) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontColour(colour);
  }

  /**
   * Sets the script style (eg. superscript, subscript) for this font, 
   * if the font hasn't been initialized
   * 
   * @param scriptStyle the colour
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setScriptStyle(int scriptStyle) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setFontScriptStyle(scriptStyle);
  }

  /** 
   * Sets the struck out flag
   *
   * @param so TRUE if the font is struck out, false otherwise
   * @exception WriteException, if this font is already in use elsewhere
   */
  protected void setStruckout(boolean os) throws WriteException
  {
    if (isInitialized())
    { 
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setFontStruckout(os);
  }
}
