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

package jxl.read.biff;

import jxl.common.Logger;

import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;

/**
 * Contains the cell dimensions of this worksheet
 */
class Window2Record extends RecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Window2Record.class);

  /**
   * Selected flag
   */
  private boolean selected;
  /**
   * Show grid lines flag
   */
  private boolean showGridLines;
  /**
   * Display zero values flag
   */
  private boolean displayZeroValues;
  /**
   * The window contains frozen panes
   */
  private boolean frozenPanes;
  /**
   * The window contains panes that are frozen but not split
   */
  private boolean frozenNotSplit;
  /**
   * The view mode:  normal or pagebreakpreview
   */
  private boolean pageBreakPreviewMode;

  /**
   * The page break preview magnification
   */
  private int pageBreakPreviewMagnification;

  /**
   * The normal view  magnification
   */
  private int normalMagnification;

  // Dummy overload
  private static class Biff7 {};
  public static final Biff7 biff7 = new Biff7();

  /**
   * Constructs the dimensions from the raw data
   *
   * @param t the raw data
   */
  public Window2Record(Record t)
  {
    super(t);
    byte[] data = t.getData();

    int options = IntegerHelper.getInt(data[0], data[1]);

    selected = ((options & 0x200) != 0);
    showGridLines = ((options & 0x02) != 0);
    frozenPanes = ((options & 0x08) != 0);
    displayZeroValues = ((options & 0x10) != 0);
    frozenNotSplit = ((options & 0x100) != 0);
    pageBreakPreviewMode = ((options & 0x800) != 0);

    pageBreakPreviewMagnification = IntegerHelper.getInt(data[10], data[11]);
    normalMagnification = IntegerHelper.getInt(data[12], data[13]);
  }

  /**
   * Constructs the dimensions from the raw data.  Dummy overload for
   * biff7 workbooks
   *
   * @param t the raw data
   * @param dummy the overload
   */
  public Window2Record(Record t, Biff7 biff7)
  {
    super(t);
    byte[] data = t.getData();

    int options = IntegerHelper.getInt(data[0], data[1]);

    selected = ((options & 0x200) != 0);
    showGridLines = ((options & 0x02) != 0);
    frozenPanes = ((options & 0x08) != 0);
    displayZeroValues = ((options & 0x10) != 0);
    frozenNotSplit = ((options & 0x100) != 0);
    pageBreakPreviewMode = ((options & 0x800) != 0);
  }

  /**
   * Accessor for the selected flag
   *
   * @return TRUE if this sheet is selected, FALSE otherwise
   */
  public boolean isSelected()
  {
    return selected;
  }

  /**
   * Accessor for the show grid lines flag
   *
   * @return TRUE to show grid lines, FALSE otherwise
   */
  public boolean getShowGridLines()
  {
    return showGridLines;
  }

  /**
   * Accessor for the zero values flag
   *
   * @return TRUE if this sheet displays zero values, FALSE otherwise
   */
  public boolean getDisplayZeroValues()
  {
    return displayZeroValues;
  }

  /**
   * Accessor for the frozen panes flag
   *
   * @return TRUE if this contains frozen panes, FALSE otherwise
   */
  public boolean getFrozen()
  {
    return frozenPanes;
  }

  /**
   * Accessor for the frozen not split flag
   *
   * @return TRUE if this contains frozen, FALSE otherwise
   */
  public boolean getFrozenNotSplit()
  {
    return frozenNotSplit;
  }

  /**
   * Accessor for the page break preview mode
   *
   * @return TRUE if this sheet is in page break preview, FALSE otherwise
   */
  public boolean isPageBreakPreview()
  {
    return pageBreakPreviewMode;
  }

  /**
   * Accessor for the page break preview magnification
   *
   * @return the cached paged break preview magnfication factor in percent
   */
  public int getPageBreakPreviewMagnificaiton()
  {
    return pageBreakPreviewMagnification;
  }

  /**
   * Accessor for the normal view  magnification
   *
   * @return the cached normal view magnfication factor in percent
   */
  public int getNormalMagnificaiton()
  {
    return normalMagnification;
  }

}







