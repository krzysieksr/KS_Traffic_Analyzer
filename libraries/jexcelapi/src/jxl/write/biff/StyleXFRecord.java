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

/**
 * A style XF Record
 */
public class StyleXFRecord extends XFRecord
{
  /**
   * Constructor
   * 
   * @param fnt the font for this style
   * @param form the format of this style
   */
  public StyleXFRecord(FontRecord fnt, DisplayFormat form)
  {
    super(fnt, form);
    
    setXFDetails(XFRecord.style, 0xfff0);
  }


  /**
   * Sets the raw cell options.  Called by WritableFormattingRecord
   * when setting the built in cell formats
   * 
   * @param opt the cell options
   */
  public final void setCellOptions(int opt)
  {
    super.setXFCellOptions(opt);
  }

  /**
   * Sets whether or not this XF record locks the cell
   * 
   * @param l the locked flag
   * @exception WriteException 
   */
  public void setLocked(boolean l)
  {
    super.setXFLocked(l);
  }

}
