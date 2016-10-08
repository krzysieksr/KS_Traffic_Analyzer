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
 * The calculation mode for the workbook, as set from the Options
 * dialog box
 */
class CalcModeRecord extends WritableRecordData
{
  /**
   * The calculation mode (manual, automatic)
   */
  private CalcMode calculationMode;

  private static class CalcMode
  {
    /**
     * The indicator as written to the output file
     */
    int value;

    /**
     * Constructor
     * 
     * @param m 
     */
    public CalcMode(int m)
    {
      value = m;
    }
  }

  /**
   * Manual calculation
   */
  static CalcMode manual = new CalcMode(0);
  /**
   * Automatic calculation
   */
  static CalcMode automatic = new CalcMode(1);
  /**
   * Automatic calculation, except tables
   */
  static CalcMode automaticNoTables  = new CalcMode(-1);
  
  /**
   * Constructor
   * 
   * @param cm the calculation mode
   */
  public CalcModeRecord(CalcMode cm)
  {
    super(Type.CALCMODE);
    calculationMode = cm;
  }


  /**
   * Gets the binary to data to write to the output file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[2];
    
    IntegerHelper.getTwoBytes(calculationMode.value, data, 0);

    return data;
  }
}


