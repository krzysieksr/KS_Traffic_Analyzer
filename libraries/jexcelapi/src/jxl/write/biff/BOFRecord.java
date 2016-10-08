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

import jxl.biff.Type;
import jxl.biff.WritableRecordData;

/**
 * Record to indicate the beginning of a new stream in the Compound
 * File
 */
class BOFRecord extends WritableRecordData
{
  /**
   * The data to write to the file
   */
  private byte[] data;

  // Dummy types for constructor overloading
  private static class WorkbookGlobalsBOF{};
  private static class SheetBOF{};

  public final static WorkbookGlobalsBOF workbookGlobals
    = new WorkbookGlobalsBOF();
  public final static SheetBOF sheet = new SheetBOF();

  /**
   * Constructor for generating a workbook globals BOF record
   *
   * @param dummy - a dummy argument for overloading purposes
   */
  public BOFRecord(WorkbookGlobalsBOF dummy)
  {
    super(Type.BOF);

    // Create the data as biff 8 format with a substream type of 
    // workbook globals
    data = new byte[] 
    {(byte) 0x0, 
     (byte) 0x6, 
     (byte) 0x5, // substream type 
     (byte) 0x0, // substream type
     (byte) 0xf2, // rupBuild 
     (byte) 0x15, // rupBuild
     (byte) 0xcc, // rupYear
     (byte) 0x07, // rupYear
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x6, // sfo
     (byte) 0x0, // sfo,
     (byte) 0x0, // sfo
     (byte) 0x0  // sfo
    };
  }

  /**
   * Constructor for generating a sheet BOF record
   *
   * @param dummy - a dummy argument for overloading purposes
   */
  public BOFRecord(SheetBOF dummy)
  {
    super(Type.BOF);

    // Create the data as biff 8 format with a substream type of 
    // worksheet
    data = new byte[] 
    {(byte) 0x0, 
     (byte) 0x6, 
     (byte) 0x10, // substream type 
     (byte) 0x0,  // substream type
     (byte) 0xf2, // rupBuild 
     (byte) 0x15, // rupBuild
     (byte) 0xcc, // rupYear
     (byte) 0x07, // rupYear
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x0, // bfh
     (byte) 0x6, // sfo
     (byte) 0x0, // sfo,
     (byte) 0x0, // sfo
     (byte) 0x0  // sfo
    };
  }

  /**
   * Gets the data for writing to the output file
   *
   * @return the binary data for writing
   */
  public byte[] getData()
  {
    return data;
  }
}

