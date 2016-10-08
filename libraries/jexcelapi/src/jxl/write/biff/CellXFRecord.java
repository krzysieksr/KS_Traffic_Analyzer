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

import jxl.biff.DisplayFormat;
import jxl.biff.FontRecord;
import jxl.biff.XFRecord;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Orientation;
import jxl.format.Pattern;
import jxl.format.VerticalAlignment;
import jxl.write.WriteException;

/**
 * A cell XF Record
 */
public class CellXFRecord extends XFRecord
{
  /**
   * Constructor
   * 
   * @param fnt the font
   * @param form the format
   */
  protected CellXFRecord(FontRecord fnt, DisplayFormat form )
  {
    super(fnt, form);
    setXFDetails(XFRecord.cell,0);
  }

  /**
   * Copy constructor.  Invoked when copying formats to handle cell merging
   * 
   * @param fmt the format to copy
   */
  CellXFRecord(XFRecord fmt )
  {
    super(fmt);
    setXFDetails(XFRecord.cell,0);
  }

  /**
   * A public copy constructor which can be used for copy formats between
   * different sheets
   */
  protected CellXFRecord(CellFormat format)
  {
    super(format);
  }

  /**
   * Sets the alignment for the cell
   * 
   * @exception WriteException 
   * @param a the alignment
   */
  public void setAlignment(Alignment a) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setXFAlignment(a);
  }

  /**
   * Sets the background for the cell
   * 
   * @exception WriteException 
   * @param c the background colour
   * @param p the background patter
   */
  public void setBackground(Colour c, Pattern p) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setXFBackground(c, p);
    super.setXFCellOptions(0x4000);
  }

  /**
   * Sets whether or not this XF record locks the cell
   * 
   * @param l the locked flag
   * @exception WriteException 
   */
  public void setLocked(boolean l) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setXFLocked(l);
    super.setXFCellOptions(0x8000);
  }

  /**
   * Sets the indentation of the cell text
   *
   * @param i the indentation
   */
  public void setIndentation(int i) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setXFIndentation(i);
  }

  /**
   * Sets the shrink to fit flag
   *
   * @param b the shrink to fit flag
   */
  public void setShrinkToFit(boolean s) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }
    super.setXFShrinkToFit(s);
  }

  /**
   * Sets the vertical alignment for cells with this style
   * 
   * @exception WriteException 
   * @param va the vertical alignment
   */
  public void setVerticalAlignment(VerticalAlignment va)
    throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setXFVerticalAlignment(va);
  }

  /**
   * Sets the text orientation for cells with this style
   * 
   * @exception WriteException 
   * @param o the orientation
   */
  public void setOrientation(Orientation o)
    throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setXFOrientation(o);
  }

  /**
   * Sets the text wrapping for cells with this style.  If the parameter is
   * set to TRUE, then data in this cell will be wrapped around, and the
   * cell's height adjusted accordingly
   * 
   * @exception WriteException 
   * @param w the wrap
   */
  public void setWrap(boolean w) throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    super.setXFWrap(w);
  }

  /**
   * Sets the border style for cells with this format
   * 
   * @exception WriteException 
   * @param b the border
   * @param ls the line for the specified border
   */
  public void setBorder(Border b, BorderLineStyle ls, Colour c) 
    throws WriteException
  {
    if (isInitialized())
    {
      throw new JxlWriteException(JxlWriteException.formatInitialized);
    }

    if (b == Border.ALL)
    {
      // Apply to all
      super.setXFBorder(Border.LEFT, ls, c);
      super.setXFBorder(Border.RIGHT, ls, c);
      super.setXFBorder(Border.TOP, ls, c);
      super.setXFBorder(Border.BOTTOM, ls, c);
      return;
    }

    if (b == Border.NONE)
    {
      // Apply to all
      super.setXFBorder(Border.LEFT,   BorderLineStyle.NONE, Colour.BLACK);
      super.setXFBorder(Border.RIGHT,  BorderLineStyle.NONE, Colour.BLACK);
      super.setXFBorder(Border.TOP,    BorderLineStyle.NONE, Colour.BLACK);
      super.setXFBorder(Border.BOTTOM, BorderLineStyle.NONE, Colour.BLACK);
      return;
    }

    super.setXFBorder(b, ls, c);
  }
}

