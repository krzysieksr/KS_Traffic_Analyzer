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

import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Record which stores the maximum iterations option from the Options
 * dialog box
 */
class CalcCountRecord extends WritableRecordData
{
  /**
   * The iteration count
   */
  private int calcCount;
  /**
   * The binary data to write to the output file
   */
  private  byte[] data;
  
  /**
   * Constructor
   * 
   * @param cnt the count indicator
   */
  public CalcCountRecord(int cnt)
  {
    super(Type.CALCCOUNT);
    calcCount = cnt;
  }


  /**
   * Gets the data to write out to the file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[2];
    
    IntegerHelper.getTwoBytes(calcCount, data, 0);

    return data;
  }
}


