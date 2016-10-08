/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

import jxl.common.Logger;

import jxl.biff.XFRecord;
import jxl.write.DateFormat;
import jxl.write.DateFormats;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableWorkbook;

/**
 * A structure containing the styles used by this workbook.  This is used
 * to enforce thread safety by tying the default styles to a workbook
 * instance rather than by initializing them statically
 */
class Styles
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Styles.class);

  /**
   * The default font for Cell formats
   */
  private WritableFont  arial10pt;

  /**
   * The font used for hyperlinks
   */
  private WritableFont hyperlinkFont;

  /**
   * The default style for cells
   */
  private WritableCellFormat normalStyle;

  /**
   * The style used for hyperlinks
   */
  private WritableCellFormat hyperlinkStyle;

  /**
   * A cell format used to hide the cell contents
   */
  private WritableCellFormat hiddenStyle;

  /**
   * A cell format used for the default date format
   */
  private WritableCellFormat defaultDateFormat;

  /**
   * Constructor
   */
  public Styles()
  {
    arial10pt = null;
    hyperlinkFont = null;
    normalStyle = null;
    hyperlinkStyle = null;
    hiddenStyle = null;
  }

  private synchronized void initNormalStyle()
  {
    normalStyle = new WritableCellFormat(getArial10Pt(), 
                                         NumberFormats.DEFAULT);
    normalStyle.setFont(getArial10Pt());
  }

  public WritableCellFormat getNormalStyle()
  {
    if (normalStyle == null)
    {
      initNormalStyle();
    }

    return normalStyle;
  }

  private synchronized void initHiddenStyle()
  {
    hiddenStyle = new WritableCellFormat
      (getArial10Pt(), new DateFormat(";;;"));
  }

  public WritableCellFormat getHiddenStyle()
  {
    if (hiddenStyle == null)
    {
      initHiddenStyle();
    }

    return hiddenStyle;
  }

  private synchronized void initHyperlinkStyle()
  {
    hyperlinkStyle = new WritableCellFormat(getHyperlinkFont(), 
                                            NumberFormats.DEFAULT);
  }

  public WritableCellFormat getHyperlinkStyle()
  {
    if (hyperlinkStyle == null)
    {
      initHyperlinkStyle();
    }

    return hyperlinkStyle;
  }

  private synchronized void initArial10Pt()
  {
    arial10pt = new WritableFont(WritableWorkbook.ARIAL_10_PT);
  }

  public WritableFont getArial10Pt()
  {
    if (arial10pt == null)
    {
      initArial10Pt();
    }

    return arial10pt;
  }

  private synchronized void initHyperlinkFont()
  {
    hyperlinkFont = new WritableFont(WritableWorkbook.HYPERLINK_FONT);
  }

  public WritableFont getHyperlinkFont()
  {
    if (hyperlinkFont == null)
    {
      initHyperlinkFont();
    }

    return hyperlinkFont;
  }

  private synchronized void initDefaultDateFormat()
  {
    defaultDateFormat = new WritableCellFormat(DateFormats.DEFAULT);
  }

  public WritableCellFormat getDefaultDateFormat()
  {
    if (defaultDateFormat == null)
    {
      initDefaultDateFormat();
    }

    return defaultDateFormat;
  }

  /**
   * Gets the thread safe version of the cell format passed in.  If the 
   * format is already thread safe (ie. it doesn't use a statically initialized
   * format or font) then the same object is simply returned
   * This object is already tied to a workbook instance, so no synchronisation
   * is necesasry
   *
   * @param wf a format to verify
   * @return the thread safe format
   */
  public XFRecord getFormat(XFRecord wf)
  {
    XFRecord format = wf;

    // Check to see if the format is one of the shared Workbook defaults.  If
    // so, then get hold of the Workbook's specific instance
    if (format == WritableWorkbook.NORMAL_STYLE)
    {
      format = getNormalStyle();
    }
    else if (format == WritableWorkbook.HYPERLINK_STYLE)
    {
      format = getHyperlinkStyle();
    }
    else if (format == WritableWorkbook.HIDDEN_STYLE)
    {
      format = getHiddenStyle();
    }
    else if (format == DateRecord.defaultDateFormat)
    {
      format = getDefaultDateFormat();
    }

    // Do the same with the statically shared fonts
    if (format.getFont() == WritableWorkbook.ARIAL_10_PT)
    {
      format.setFont(getArial10Pt());
    }
    else if (format.getFont() == WritableWorkbook.HYPERLINK_FONT)
    {
      format.setFont(getHyperlinkFont());
    }

    return format;
  }
}
