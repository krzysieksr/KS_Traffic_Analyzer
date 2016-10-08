/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import java.io.IOException;

import jxl.write.biff.File;

/**
 * Information for autofiltering
 */
public class AutoFilter
{
  private FilterModeRecord filterMode;
  private AutoFilterInfoRecord autoFilterInfo;
  private AutoFilterRecord autoFilter;

  /**
   * Constructor
   */
  public AutoFilter(FilterModeRecord fmr, 
                    AutoFilterInfoRecord afir)
  {
    filterMode = fmr;
    autoFilterInfo = afir;
  }

  public void add(AutoFilterRecord af)
  {
    autoFilter = af; // make this into a list sometime
  }

  /**
   * Writes out the data validation
   * 
   * @exception IOException 
   * @param outputFile the output file
   */
  public void write(File outputFile) throws IOException
  {
    if (filterMode != null)
    {
      outputFile.write(filterMode);
    }

    if (autoFilterInfo != null)
    {
      outputFile.write(autoFilterInfo);
    }
    
    if (autoFilter != null)
    {
      outputFile.write(autoFilter);
    }
  }
}
