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

package jxl.write;

import java.io.File;
import java.net.URL;

import jxl.Hyperlink;
import jxl.write.biff.HyperlinkRecord;

/**
 * A writable hyperlink.  Provides API to modify the contents of the hyperlink
 */
public class WritableHyperlink extends HyperlinkRecord implements Hyperlink
{
  /**
   * Constructor used internally by the worksheet when making a copy
   * of worksheet
   *
   * @param h the hyperlink being read in
   * @param ws the writable sheet containing the hyperlink
   */
  public WritableHyperlink(Hyperlink h, WritableSheet ws)
  {
    super(h, ws);
  }

  /**
   * Constructs a URL hyperlink in a single cell
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param url the hyperlink
   */
  public WritableHyperlink(int col, int row, URL url)
  {
    this(col, row, col, row, url);
  }

  /**
   * Constructs a url hyperlink to a range of cells
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param url the hyperlink
   */
  public WritableHyperlink(int col, int row, int lastcol, int lastrow, URL url)
  {
    this(col, row, lastcol, lastrow, url, null);
  }

  /**
   * Constructs a url hyperlink to a range of cells
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param url the hyperlink
   * @param desc the description text to place in the cell
   */
  public WritableHyperlink(int col,
                           int row,
                           int lastcol,
                           int lastrow,
                           URL url,
                           String desc)
  {
    super(col, row, lastcol, lastrow, url, desc);
  }

  /**
   * Constructs a file hyperlink in a single cell
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param file the hyperlink
   */
  public WritableHyperlink(int col, int row, File file)
  {
    this(col, row, col, row, file, null);
  }

  /**
   * Constructs a file hyperlink in a single cell
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param file the hyperlink
   * @param desc the hyperlink description
   */
  public WritableHyperlink(int col, int row, File file, String desc)
  {
    this(col, row, col, row, file, desc);
  }

  /**
   * Constructs a File hyperlink to a range of cells
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param file the hyperlink
   */
  public WritableHyperlink(int col, int row, int lastcol, int lastrow,
                           File file)
  {
    super(col, row, lastcol, lastrow, file, null);
  }

  /**
   * Constructs a File hyperlink to a range of cells
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param file the hyperlink
   * @param desc the description
   */
  public WritableHyperlink(int col,
                           int row,
                           int lastcol,
                           int lastrow,
                           File file,
                           String desc)
  {
    super(col, row, lastcol, lastrow, file, desc);
  }

  /**
   * Constructs a hyperlink to some cells within this workbook
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param desc the cell contents for this hyperlink
   * @param sheet the sheet containing the cells to be linked to
   * @param destcol the column number of the first destination linked cell
   * @param destrow the row number of the first destination linked cell
   */
  public WritableHyperlink(int col, int row,
                           String desc,
                           WritableSheet sheet,
                           int destcol, int destrow)
  {
    this(col, row, col, row,
         desc,
         sheet, destcol, destrow, destcol, destrow);
  }

  /**
   * Constructs a hyperlink to some cells within this workbook
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param desc the cell contents for this hyperlink
   * @param sheet the sheet containing the cells to be linked to
   * @param destcol the column number of the first destination linked cell
   * @param destrow the row number of the first destination linked cell
   * @param lastdestcol the column number of the last destination linked cell
   * @param lastdestrow the row number of the last destination linked cell
   */
  public WritableHyperlink(int col, int row,
                           int lastcol, int lastrow,
                           String desc,
                           WritableSheet sheet,
                           int destcol, int destrow,
                           int lastdestcol, int lastdestrow)
  {
    super(col, row, lastcol, lastrow,
          desc,
          sheet, destcol, destrow,
          lastdestcol, lastdestrow);
  }

  /**
   * Sets the URL of this hyperlink
   *
   * @param url the url
   */
  public void setURL(URL url)
  {
    super.setURL(url);
  }

  /**
   * Sets the file activated by this hyperlink
   *
   * @param file the file
   */
  public void setFile(File file)
  {
    super.setFile(file);
  }

  /**
   * Sets the description to appear in the hyperlink cell
   *
   * @param desc the description
   */
  public void setDescription(String desc)
  {
    super.setContents(desc);
  }

  /**
   * Sets the location of the cells to be linked to within this workbook
   *
   * @param desc the label describing the link
   * @param sheet the sheet containing the cells to be linked to
   * @param destcol the column number of the first destination linked cell
   * @param destrow the row number of the first destination linked cell
   * @param lastdestcol the column number of the last destination linked cell
   * @param lastdestrow the row number of the last destination linked cell
   */
  public void setLocation(String desc,
                          WritableSheet sheet,
                          int destcol, int destrow,
                          int lastdestcol, int lastdestrow)
  {
    super.setLocation(desc, sheet, destcol, destrow, lastdestcol, lastdestrow);
  }

}


