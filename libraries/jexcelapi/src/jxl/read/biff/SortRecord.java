/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan, Al Mantei
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

import jxl.biff.RecordData;
import jxl.biff.StringHelper;
import jxl.biff.Type;

/**
 * A storage area for the last Sort dialog box area
 */
public class SortRecord extends RecordData 
{
	private int col1Size;
	private int col2Size;
	private int col3Size;
	private String col1Name;
	private String col2Name;
	private String col3Name;
	private byte optionFlags;
	private boolean sortColumns = false;
	private boolean sortKey1Desc = false;
	private boolean sortKey2Desc = false;
	private boolean sortKey3Desc = false;
	private boolean sortCaseSensitive = false;
	/**
	 * Constructs this object from the raw data
	 *
	 * @param r the raw data
	 */
	public SortRecord(Record r) 
  {
		super(Type.SORT);

		byte[] data = r.getData();

		optionFlags = data[0];

		sortColumns = ((optionFlags & 0x01) != 0);
		sortKey1Desc = ((optionFlags & 0x02) != 0);
		sortKey2Desc = ((optionFlags & 0x04) != 0);
		sortKey3Desc = ((optionFlags & 0x08) != 0);
		sortCaseSensitive = ((optionFlags & 0x10) != 0);

		// data[1] contains sort list index - not implemented...

		col1Size = data[2];
		col2Size = data[3];
		col3Size = data[4];
		int curPos = 5;
		if (data[curPos++] == 0x00) 
    {
			col1Name = new String(data, curPos, col1Size);
			curPos += col1Size;
		}
    else 
    {
			col1Name = StringHelper.getUnicodeString(data, col1Size, curPos);
			curPos += col1Size * 2;
		}

		if (col2Size > 0) 
    {
			if (data[curPos++] == 0x00) 
      {
				col2Name = new String(data, curPos, col2Size);
				curPos += col2Size;
			} 
      else 
      {
				col2Name = StringHelper.getUnicodeString(data, col2Size, curPos);
				curPos += col2Size * 2;
			}
		} 
    else
    {
			col2Name = "";
    }
		if (col3Size > 0) 
    {
			if (data[curPos++] == 0x00) 
      {
				col3Name = new String(data, curPos, col3Size);
				curPos += col3Size;
			} 
      else 
      {
				col3Name = StringHelper.getUnicodeString(data, col3Size, curPos);
				curPos += col3Size * 2;
			}
		} 
    else
    {
			col3Name = "";
    }
	}

	/**
	 * Accessor for the 1st Sort Column Name
	 *
	 * @return the 1st Sort Column Name
	 */
	public String getSortCol1Name() {
		return col1Name;
	}
	/**
	 * Accessor for the 2nd Sort Column Name
	 *
	 * @return the 2nd Sort Column Name
	 */
	public String getSortCol2Name() {
		return col2Name;
	}
	/**
	 * Accessor for the 3rd Sort Column Name
	 *
	 * @return the 3rd Sort Column Name
	 */
	public String getSortCol3Name() {
		return col3Name;
	}
	/**
	 * Accessor for the Sort by Columns flag
	 *
	 * @return the Sort by Columns flag
	 */
	public boolean getSortColumns() {
		return sortColumns;
	}
	/**
	 * Accessor for the Sort Column 1 Descending flag
	 *
	 * @return the Sort Column 1 Descending flag
	 */
	public boolean getSortKey1Desc() {
		return sortKey1Desc;
	}
	/**
	 * Accessor for the Sort Column 2 Descending flag
	 *
	 * @return the Sort Column 2 Descending flag
	 */
	public boolean getSortKey2Desc() {
		return sortKey2Desc;
	}
	/**
	 * Accessor for the Sort Column 3 Descending flag
	 *
	 * @return the Sort Column 3 Descending flag
	 */
	public boolean getSortKey3Desc() {
		return sortKey3Desc;
	}
	/**
	 * Accessor for the Sort Case Sensitivity flag
	 *
	 * @return the Sort Case Secsitivity flag
	 */
	public boolean getSortCaseSensitive() {
		return sortCaseSensitive;
	}
}
