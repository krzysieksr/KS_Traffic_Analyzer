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

import jxl.common.Logger;

import jxl.SheetSettings;
import jxl.biff.DoubleHelper;
import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.format.PageOrder;
import jxl.format.PageOrientation;
import jxl.format.PaperSize;

/**
 * Stores the options and measurements from the Page Setup dialog box
 */
class SetupRecord extends WritableRecordData
{
  /**
   * The logger
   */
  Logger logger = Logger.getLogger(SetupRecord.class);

  /**
   * The binary data for output to file
   */
  private byte[] data;

  /**
   * The header margin
   */
  private double          headerMargin;

  /**
   * The footer margin
   */
  private double          footerMargin;

  /**
   * The page orientation
   */
  private PageOrientation orientation;

  /**
   * The page order
   */
  private PageOrder order;

  /**
   * The paper size
   */
  private int             paperSize;

  /**
   * The scale factor
   */
  private int             scaleFactor;

  /**
   * The page start
   */
  private int             pageStart;

  /**
   * The fit width
   */
  private int             fitWidth;

  /**
   * The fit height
   */
  private int             fitHeight;

  /**
   * The horizontal print resolution
   */
  private int         horizontalPrintResolution;

  /**
   * The vertical print resolution
   */
  private int         verticalPrintResolution;    

  /**
   * The number of copies
   */
  private int         copies;

  /**
   * Indicates whether the setup data should be  initiliazed in the setup
   * box
   */
  private boolean initialized;

  /**
   * Constructor, taking the sheet settings.  This object just
   * takes the various fields from the bean in which it is interested
   * 
   * @param the sheet settings
   */
  public SetupRecord(SheetSettings s)
  {
    super(Type.SETUP);

    orientation = s.getOrientation();
    order = s.getPageOrder();
    headerMargin = s.getHeaderMargin();
    footerMargin = s.getFooterMargin();
    paperSize = s.getPaperSize().getValue();
    horizontalPrintResolution = s.getHorizontalPrintResolution();
    verticalPrintResolution = s.getVerticalPrintResolution();
    fitWidth = s.getFitWidth();
    fitHeight = s.getFitHeight();
    pageStart = s.getPageStart();
    scaleFactor = s.getScaleFactor();
    copies = s.getCopies();
    initialized = true;
  }

  /**
   * Sets the orientation
   *
   * @param o the orientation
   */
  public void setOrientation(PageOrientation o)
  {
    orientation = o;
  }

  /**
   * Sets the page order
   * 
   * @param o
   */
  public void setOrder(PageOrder o)
  {
    order = o;
  }

  /**
   * Sets the header and footer margins
   * 
   * @param hm the header margin
   * @param fm the footer margin
   */
  public void setMargins(double hm, double fm)
  {
    headerMargin = hm;
    footerMargin = fm;
  }

  /**
   * Sets the paper size
   *
   * @param ps the paper size
   */
  public void setPaperSize(PaperSize ps)
  {
    paperSize = ps.getValue();
  }

  /**
   * Gets the binary data for output to file
   *
   * @return the binary data
   */
  public byte[] getData()
  {
    data = new byte[34];

    // Paper size
    IntegerHelper.getTwoBytes(paperSize, data, 0);

    // Scale factor
    IntegerHelper.getTwoBytes(scaleFactor, data, 2);

    // Page start
    IntegerHelper.getTwoBytes(pageStart, data, 4);

    // Fit width
    IntegerHelper.getTwoBytes(fitWidth, data, 6);

    // Fit height
    IntegerHelper.getTwoBytes(fitHeight, data, 8);

    // grbit
    int options = 0;
    if (order == PageOrder.RIGHT_THEN_DOWN)
    {
      options |= 0x01;
    }

    if (orientation == PageOrientation.PORTRAIT)
    {
      options |= 0x02;
    }

    if (pageStart != 0)
    {
      options |= 0x80;
    }

    if (!initialized)
    {
      options |= 0x04;
    }

    IntegerHelper.getTwoBytes(options, data, 10);

    // print resolution
    IntegerHelper.getTwoBytes(horizontalPrintResolution, data, 12);

    // vertical print resolution
    IntegerHelper.getTwoBytes(verticalPrintResolution, data, 14);
    
    // header margin
    DoubleHelper.getIEEEBytes(headerMargin, data, 16);

    // footer margin
    DoubleHelper.getIEEEBytes(footerMargin, data, 24);

    // Number of copies
    IntegerHelper.getTwoBytes(copies, data, 32);

    return data;
  }
}

