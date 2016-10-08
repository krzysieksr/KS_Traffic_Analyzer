/*********************************************************************
*
*      Copyright (C) 2001 Andrew Khan
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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.CellType;
import jxl.LabelCell;
import jxl.biff.FormattingRecords;
import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.format.CellFormat;

/**
 * A label record, used for writing out string
 */
public abstract class LabelRecord extends CellValue
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(LabelRecord.class);  

  /**
   * The string
   */
  private String contents;

  /**
   * A handle to the shared strings used within this workbook
   */
  private SharedStrings sharedStrings;

  /**
   * The index of the string in the shared string table
   */
  private int index;

  /**
   * Constructor used when creating a label from the user API
   * 
   * @param c the column
   * @param cont the contents
   * @param r the row
   */
  protected LabelRecord(int c, int r, String cont)
  {
    super(Type.LABELSST, c, r);
    contents = cont;
    if (contents == null)
    {
      contents="";
    }
  }

  /**
   * Constructor used when creating a label from the API.  This is 
   * overloaded to allow formatting information to be passed to the record
   * 
   * @param c the column
   * @param cont the contents
   * @param r the row
   * @param st the format applied to the cell
   */
  protected LabelRecord(int c, int r, String cont, CellFormat st)
  {
    super(Type.LABELSST, c, r, st);
    contents = cont;

    if (contents == null)
    {
      contents="";
    }
  }


  /**
   * Copy constructor
   * 
   * @param c the column
   * @param r the row
   * @param nr the record to copy
   */
  protected LabelRecord(int c, int r, LabelRecord lr)
  {
    super(Type.LABELSST, c, r, lr);
    contents = lr.contents;
  }

  /**
   * Constructor used when copying a label from a read only
   * spreadsheet
   * 
   * @param lc the label to copy
   */
  protected LabelRecord(LabelCell lc)
  {
    super(Type.LABELSST, lc);
    contents = lc.getString();
    if (contents == null)
    {
      contents = "";
    }
  }

  /**
   * Returns the content type of this cell
   * 
   * @return the content type for this cell
   */
  public CellType getType()
  {
    return CellType.LABEL;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] celldata = super.getData();
    byte[] data = new byte[celldata.length + 4];
    System.arraycopy(celldata, 0, data, 0, celldata.length);
    IntegerHelper.getFourBytes(index, data, celldata.length);

    return data;
  }

  /**
   * Quick and dirty function to return the contents of this cell as a string.
   * For more complex manipulation of the contents, it is necessary to cast
   * this interface to correct subinterface
   * 
   * @return the contents of this cell as a string
   */
  public String getContents()
  {
    return contents;
  }

  /**
   * Gets the label for this cell.  The value returned will be the same
   * as for the getContents method in the base class
   * 
   * @return the cell contents
   */
  public String getString()
  {
    return contents;
  }

  /**
   * Sets the string contents of this cell
   * 
   * @param s the new string contents
   */
  protected void setString(String s)
  {
    if (s == null)
    {
      s = "";
    }

    contents = s;

    // Don't bother doing anything if this cell has not been referenced
    // yet - everything will be set up in due course
    if (!isReferenced())
    {
      return;
    }

    Assert.verify(sharedStrings != null);

    // Initalize the shared string index
    index = sharedStrings.getIndex(contents);

    // Use the sharedStrings reference instead of this object's own
    // handle - this means that the bespoke copy becomes eligible for
    // garbage collection
    contents = sharedStrings.get(index);
  }

  /**
   * Overrides the method in the base class in order to add the string
   * content to the shared string table, and to store its shared string
   * index
   *
   * @param fr the formatting records
   * @param ss the shared strings used within the workbook
   * @param s
   */
  void setCellDetails(FormattingRecords fr, SharedStrings ss, 
                      WritableSheetImpl s)
  {
    super.setCellDetails(fr, ss, s);

    sharedStrings = ss;

    index = sharedStrings.getIndex(contents);

    // Use the sharedStrings reference instead of this object's own
    // handle - this means that the bespoke copy becomes eligible for
    // garbage collection
    contents = sharedStrings.get(index);
  }

}



