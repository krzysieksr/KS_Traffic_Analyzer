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

import java.io.File;
import java.net.URL;
import java.net.MalformedURLException;
import java.util.ArrayList;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.CellType;
import jxl.Hyperlink;
import jxl.Range;
import jxl.WorkbookSettings;
import jxl.biff.CellReferenceHelper;
import jxl.biff.IntegerHelper;
import jxl.biff.SheetRangeImpl;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.write.Label;
import jxl.write.WritableHyperlink;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;


/**
 * A hyperlink
 */
public class HyperlinkRecord extends WritableRecordData
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(HyperlinkRecord.class);

  /**
   * The first row
   */
  private int firstRow;
  /**
   * The last row
   */
  private int lastRow;
  /**
   * The first column
   */
  private int firstColumn;
  /**
   * The last column
   */
  private int lastColumn;

  /**
   * The URL referred to by this hyperlink
   */
  private URL url;

  /**
   * The local file referred to by this hyperlink
   */
  private File file;

  /**
   * The location in this workbook referred to by this hyperlink
   */
  private String location;

  /**
   * The cell contents of the cell which activate this hyperlink
   */
  private String contents;

  /**
   * The type of this hyperlink
   */
  private LinkType linkType;

  /**
   * The data for this hyperlink
   */
  private byte[] data;

  /**
   * The range of this hyperlink.  When creating a hyperlink, this will
   * be null until the hyperlink is added to the sheet
   */
  private Range range;

  /**
   * The sheet containing this hyperlink
   */
  private WritableSheet sheet;

  /**
   * Indicates whether this record has been modified since it was copied
   */
  private boolean modified;

  /**
   * The excel type of hyperlink
   */
  private static class LinkType {};

  private static final LinkType urlLink      = new LinkType();
  private static final LinkType fileLink     = new LinkType();
  private static final LinkType uncLink      = new LinkType();
  private static final LinkType workbookLink = new LinkType();
  private static final LinkType unknown      = new LinkType();

  /**
   * Constructs this object from the readable spreadsheet
   *
   * @param hl the hyperlink from the read spreadsheet
   */
  protected HyperlinkRecord(Hyperlink h, WritableSheet s)
  {
    super(Type.HLINK);

    if (h instanceof jxl.read.biff.HyperlinkRecord)
    {
      copyReadHyperlink(h, s);
    }
    else
    {
      copyWritableHyperlink(h, s);
    }
  }

  /**
   * Copies a hyperlink read in from a read only sheet
   */
  private void copyReadHyperlink(Hyperlink h, WritableSheet s)
  {
    jxl.read.biff.HyperlinkRecord hl = (jxl.read.biff.HyperlinkRecord) h;

    data = hl.getRecord().getData();
    sheet = s;

    // Populate this hyperlink with the copied data
    firstRow    = hl.getRow();
    firstColumn = hl.getColumn();
    lastRow     = hl.getLastRow();
    lastColumn  = hl.getLastColumn();
    range       = new SheetRangeImpl(s, 
                                     firstColumn, firstRow,
                                     lastColumn, lastRow);

    linkType = unknown;

    if (hl.isFile())
    {
      linkType = fileLink;
      file = hl.getFile();
    }
    else if (hl.isURL())
    {
      linkType = urlLink;
      url = hl.getURL();
    }
    else if (hl.isLocation())
    {
      linkType = workbookLink;
      location = hl.getLocation();
    }

    modified = false;
  }

  /**
   * Copies a hyperlink read in from a writable sheet.
   * Used when copying writable sheets
   *
   * @param hl the hyperlink from the read spreadsheet
   */
  private void copyWritableHyperlink(Hyperlink hl, WritableSheet s)
  {
    HyperlinkRecord h = (HyperlinkRecord) hl;

    firstRow = h.firstRow;
    lastRow = h.lastRow;
    firstColumn = h.firstColumn;
    lastColumn = h.lastColumn;

    if (h.url != null)
    {
      try
      {
        url = new URL(h.url.toString());
      }
      catch (MalformedURLException e)
      {
        // should never get a malformed url as a result url.toString()
        Assert.verify(false);
      }
    }
    
    if (h.file != null)
    {
      file = new File(h.file.getPath());
    }

    location = h.location;
    contents = h.contents;
    linkType = h.linkType;
    modified = true;

    sheet = s;
    range = new SheetRangeImpl(s, 
                               firstColumn, firstRow,
                               lastColumn, lastRow);
  }

  /**
   * Constructs a URL hyperlink to a range of cells
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param url the hyperlink
   * @param desc the description
   */
  protected HyperlinkRecord(int col, int row, 
                            int lastcol, int lastrow, 
                            URL url,
                            String desc)
  {
    super(Type.HLINK);

    firstColumn = col;
    firstRow = row;

    lastColumn = Math.max(firstColumn, lastcol);
    lastRow = Math.max(firstRow, lastrow);
    
    this.url = url;
    contents = desc;

    linkType = urlLink;

    modified = true;
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
  protected HyperlinkRecord(int col, int row, int lastcol, int lastrow, 
                            File file, String desc)
  {
    super(Type.HLINK);

    firstColumn = col;
    firstRow = row;

    lastColumn = Math.max(firstColumn, lastcol);
    lastRow = Math.max(firstRow, lastrow);
    contents = desc;

    this.file = file;

    if (file.getPath().startsWith("\\\\"))
    {
      linkType = uncLink;
    }
    else
    {
      linkType = fileLink;
    }

    modified = true;
  }

  /**
   * Constructs a hyperlink to some cells within this workbook
   *
   * @param col the column containing this hyperlink
   * @param row the row containing this hyperlink
   * @param lastcol the last column which activates this hyperlink
   * @param lastrow the last row which activates this hyperlink
   * @param desc the contents of the cell which describe this hyperlink
   * @param sheet the sheet containing the cells to be linked to
   * @param destcol the column number of the first destination linked cell
   * @param destrow the row number of the first destination linked cell
   * @param lastdestcol the column number of the last destination linked cell
   * @param lastdestrow the row number of the last destination linked cell
   */
  protected HyperlinkRecord(int col, int row,
                            int lastcol, int lastrow, 
                            String desc,
                            WritableSheet s,
                            int destcol, int destrow,
                            int lastdestcol, int lastdestrow)
  {
    super(Type.HLINK);

    firstColumn = col;
    firstRow = row;

    lastColumn = Math.max(firstColumn, lastcol);
    lastRow = Math.max(firstRow, lastrow);
    
    setLocation(s, destcol, destrow, lastdestcol, lastdestrow);
    contents = desc;

    linkType = workbookLink;

    modified = true;
  }

  /**
   * Determines whether this is a hyperlink to a file
   * 
   * @return TRUE if this is a hyperlink to a file, FALSE otherwise
   */
  public boolean isFile()
  {
    return linkType == fileLink;
  }

  /**
   * Determines whether this is a hyperlink to a UNC
   * 
   * @return TRUE if this is a hyperlink to a UNC, FALSE otherwise
   */
  public boolean isUNC()
  {
    return linkType == uncLink;
  }

  /**
   * Determines whether this is a hyperlink to a web resource
   *
   * @return TRUE if this is a URL
   */
  public boolean isURL()
  {
    return linkType == urlLink;
  }

  /**
   * Determines whether this is a hyperlink to a location in this workbook
   *
   * @return TRUE if this is a link to an internal location
   */
  public boolean isLocation()
  {
    return linkType == workbookLink;
  }

  /**
   * Returns the row number of the top left cell
   * 
   * @return the row number of this cell
   */
  public int getRow()
  {
    return firstRow;
  }

  /**
   * Returns the column number of the top left cell
   * 
   * @return the column number of this cell
   */
  public int getColumn()
  {
    return firstColumn;
  }

  /**
   * Returns the row number of the bottom right cell
   * 
   * @return the row number of this cell
   */
  public int getLastRow()
  {
    return lastRow;
  }

  /**
   * Returns the column number of the bottom right cell
   * 
   * @return the column number of this cell
   */
  public int getLastColumn()
  {
    return lastColumn;
  }

  /**
   * Gets the URL referenced by this Hyperlink
   *
   * @return the URL, or NULL if this hyperlink is not a URL
   */
  public URL getURL()
  {
    return url;
  }

  /**
   * Returns the local file eferenced by this Hyperlink
   *
   * @return the file, or NULL if this hyperlink is not a file
   */
  public File getFile()
  {
    return file;
  }

  /**
   * Gets the binary data to be written to the output file
   * 
   * @return the data to write to file
   */
  public byte[] getData()
  {
    if (!modified)
    {
      return data;
    }

    // Build up the jxl.common.data
    byte[] commonData = new byte[32];

    // Set the range of cells this hyperlink applies to
    IntegerHelper.getTwoBytes(firstRow, commonData, 0);
    IntegerHelper.getTwoBytes(lastRow, commonData, 2);
    IntegerHelper.getTwoBytes(firstColumn, commonData, 4);
    IntegerHelper.getTwoBytes(lastColumn, commonData, 6);

    // Some inexplicable byte sequence
    commonData[8]  = (byte) 0xd0;
    commonData[9]  = (byte) 0xc9;
    commonData[10] = (byte) 0xea;
    commonData[11] = (byte) 0x79;
    commonData[12] = (byte) 0xf9;
    commonData[13] = (byte) 0xba;
    commonData[14] = (byte) 0xce;
    commonData[15] = (byte) 0x11;
    commonData[16] = (byte) 0x8c;
    commonData[17] = (byte) 0x82;
    commonData[18] = (byte) 0x0;
    commonData[19] = (byte) 0xaa;
    commonData[20] = (byte) 0x0;
    commonData[21] = (byte) 0x4b;
    commonData[22] = (byte) 0xa9;
    commonData[23] = (byte) 0x0b;
    commonData[24] = (byte) 0x2;
    commonData[25] = (byte) 0x0;
    commonData[26] = (byte) 0x0;
    commonData[27] = (byte) 0x0;

    // Set up the option flags to indicate the type of this URL.  There
    // is no description
    int optionFlags = 0;
    if (isURL())
    {
      optionFlags = 3;
      
      if (contents != null)
      {
        optionFlags |= 0x14;
      }
    }
    else if (isFile())
    {
      optionFlags = 1;

      if (contents != null)
      {
        optionFlags |= 0x14;
      }
    }
    else if (isLocation())
    {
      optionFlags = 8;
    }
    else if (isUNC())
    {
      optionFlags = 259;
    }

    IntegerHelper.getFourBytes(optionFlags, commonData, 28);

    if (isURL())
    {
      data = getURLData(commonData);
    }
    else if (isFile())
    {
      data = getFileData(commonData);
    }
    else if (isLocation())
    {
      data = getLocationData(commonData);
    }
    else if (isUNC())
    {
      data = getUNCData(commonData);
    }

    return data;
  }

  /**
   * A standard toString method
   * 
   * @return the contents of this object as a string
   */
  public String toString()
  {
    if (isFile())
    {
      return file.toString();
    }
    else if (isURL())
    {
      return url.toString();
    }
    else if (isUNC())
    {
      return file.toString();
    }
    else
    {
      return "";
    }
  }

  /**
   * Gets the range of cells which activate this hyperlink
   * The get sheet index methods will all return -1, because the
   * cells will all be present on the same sheet
   *
   * @return the range of cells which activate the hyperlink or NULL
   * if this hyperlink has not been added to the sheet
   */
  public Range getRange()
  {
    return range;
  }

  /**
   * Sets the URL of this hyperlink
   *
   * @param url the url
   */
  public void setURL(URL url)
  {
    URL prevurl = this.url;
    linkType = urlLink;
    file = null;
    location = null;
    contents = null;
    this.url = url;
    modified = true;

    if (sheet == null)
    {
      // hyperlink has not been added to the sheet yet, so simply return
      return;
    }

    // Change the label on the sheet if it was a string representation of the 
    // URL
    WritableCell wc = sheet.getWritableCell(firstColumn, firstRow);
    
    if (wc.getType() == CellType.LABEL)
    {
      Label l = (Label) wc;
      String prevurlString = prevurl.toString();
      String prevurlString2 = "";
      if (prevurlString.charAt(prevurlString.length() - 1) == '/' ||
          prevurlString.charAt(prevurlString.length() - 1) == '\\')
      {
        prevurlString2 = prevurlString.substring(0, 
                                                 prevurlString.length() - 1);
      }

      if (l.getString().equals(prevurlString) ||
          l.getString().equals(prevurlString2))
      {
        l.setString(url.toString());
      }
    }   
  }

  /**
   * Sets the file activated by this hyperlink
   * 
   * @param file the file
   */
  public void setFile(File file)
  {
    linkType = fileLink;
    url = null;
    location = null;
    contents = null;
    this.file = file;
    modified = true;

    if (sheet == null)
    {
      // hyperlink has not been added to the sheet yet, so simply return
      return;
    }

    // Change the label on the sheet
    WritableCell wc = sheet.getWritableCell(firstColumn, firstRow);
    
    Assert.verify(wc.getType() == CellType.LABEL);

    Label l = (Label) wc;
    l.setString(file.toString());
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
  protected void setLocation(String desc,
                             WritableSheet sheet,
                             int destcol, int destrow,
                             int lastdestcol, int lastdestrow)
  {
    linkType = workbookLink;
    url = null;
    file = null;
    modified = true;
    contents = desc;

    setLocation(sheet, destcol, destrow, lastdestcol, lastdestrow);

    if (sheet == null)
    {
      // hyperlink has not been added to the sheet yet, so simply return
      return;
    }

    // Change the label on the sheet
    WritableCell wc = sheet.getWritableCell(firstColumn, firstRow);
    
    Assert.verify(wc.getType() == CellType.LABEL);

    Label l = (Label) wc;
    l.setString(desc);
  }

  /**
    * Initializes the location from the data passed in
   *
   * @param sheet the sheet containing the cells to be linked to
   * @param destcol the column number of the first destination linked cell
   * @param destrow the row number of the first destination linked cell
   * @param lastdestcol the column number of the last destination linked cell
   * @param lastdestrow the row number of the last destination linked cell
   */
  private void setLocation(WritableSheet sheet,
                           int destcol, int destrow,
                           int lastdestcol, int lastdestrow)
  {
    StringBuffer sb = new StringBuffer();
    sb.append('\'');
    
    if (sheet.getName().indexOf('\'') == -1)
    {
      sb.append(sheet.getName());
    }
    else 
    {
      // sb.append(sheet.getName().replaceAll("'", "''"));

      // Can't use replaceAll as it is only 1.4 compatible, so have to
      // do this the tedious way
      String sheetName = sheet.getName();
      int pos = 0 ;
      int nextPos = sheetName.indexOf('\'', pos);

      while (nextPos != -1 && pos < sheetName.length())
      {
        sb.append(sheetName.substring(pos, nextPos));
        sb.append("''");
        pos = nextPos + 1;
        nextPos = sheetName.indexOf('\'', pos);
      }
      sb.append(sheetName.substring(pos));
    }

    sb.append('\'');    
    sb.append('!');
    
    lastdestcol = Math.max(destcol, lastdestcol);
    lastdestrow = Math.max(destrow, lastdestrow);

    CellReferenceHelper.getCellReference(destcol, destrow, sb);
    sb.append(':');
    CellReferenceHelper.getCellReference(lastdestcol, lastdestrow, sb);

    location = sb.toString();
  }

  /** 
   * A row has been inserted, so adjust the range objects accordingly
   *
   * @param r the row which has been inserted
   */
  void insertRow(int r)
  {
    // This will not be called unless the hyperlink has been added to the
    // sheet
    Assert.verify(sheet != null && range != null);

    if (r > lastRow)
    {
      return;
    }

    if (r <= firstRow)
    {
      firstRow++;
      modified = true;
    }

    if (r <= lastRow)
    {
      lastRow++;
      modified = true;
    }

    if (modified)
    {
      range  = new SheetRangeImpl(sheet, 
                                  firstColumn, firstRow,
                                  lastColumn, lastRow);
    }
  }

  /** 
   * A column has been inserted, so adjust the range objects accordingly
   *
   * @param c the column which has been inserted
   */
  void insertColumn(int c)
  {
    // This will not be called unless the hyperlink has been added to the
    // sheet
    Assert.verify(sheet != null && range != null);

    if (c > lastColumn)
    {
      return;
    }

    if (c <= firstColumn)
    {
      firstColumn++;
      modified = true;
    }

    if (c <= lastColumn)
    {
      lastColumn++;
      modified = true;
    }

    if (modified)
    {
      range  = new SheetRangeImpl(sheet, 
                                  firstColumn, firstRow,
                                  lastColumn, lastRow);
    }
  }

  /** 
   * A row has been removed, so adjust the range objects accordingly
   *
   * @param r the row which has been inserted
   */
  void removeRow(int r)
  {
    // This will not be called unless the hyperlink has been added to the
    // sheet
    Assert.verify(sheet != null && range != null);

    if (r > lastRow)
    {
      return;
    }

    if (r < firstRow)
    {
      firstRow--;
      modified = true;
    }

    if (r < lastRow)
    {
      lastRow--;
      modified = true;
    }

    if (modified)
    {
      Assert.verify(range != null);
      range  = new SheetRangeImpl(sheet, 
                                  firstColumn, firstRow,
                                  lastColumn, lastRow);
    }
  }

  /** 
   * A column has been removed, so adjust the range objects accordingly
   *
   * @param c the column which has been removed
   */
  void removeColumn(int c)
  {
    // This will not be called unless the hyperlink has been added to the
    // sheet
    Assert.verify(sheet != null && range != null);

    if (c > lastColumn)
    {
      return;
    }

    if (c < firstColumn)
    {
      firstColumn--;
      modified = true;
    }

    if (c < lastColumn)
    {
      lastColumn--;
      modified = true;
    }

    if (modified)
    {
      Assert.verify(range != null);
      range  = new SheetRangeImpl(sheet, 
                                  firstColumn, firstRow,
                                  lastColumn, lastRow);
    }
  }

  /**
   * Gets the hyperlink stream specific to a URL link
   *
   * @param cd the data jxl.common.for all types of hyperlink
   * @return the raw data for a URL hyperlink
   */
  private byte[] getURLData(byte[] cd)
  {
    String urlString = url.toString();

    int dataLength = cd.length + 20 + (urlString.length() + 1)* 2;

    if (contents != null)
    {
      dataLength += 4 + (contents.length() + 1) * 2;
    }

    byte[] d = new byte[dataLength];

    System.arraycopy(cd, 0, d, 0, cd.length);
    
    int urlPos = cd.length;

    if (contents != null)
    {
      IntegerHelper.getFourBytes(contents.length() + 1, d, urlPos);
      StringHelper.getUnicodeBytes(contents, d, urlPos + 4);
      urlPos += (contents.length() + 1) * 2 + 4;
    }
    
    // Inexplicable byte sequence
    d[urlPos]    = (byte) 0xe0;
    d[urlPos+1]  = (byte) 0xc9;
    d[urlPos+2]  = (byte) 0xea;
    d[urlPos+3]  = (byte) 0x79;
    d[urlPos+4]  = (byte) 0xf9;
    d[urlPos+5]  = (byte) 0xba;
    d[urlPos+6]  = (byte) 0xce;
    d[urlPos+7]  = (byte) 0x11;
    d[urlPos+8]  = (byte) 0x8c;
    d[urlPos+9]  = (byte) 0x82;
    d[urlPos+10] = (byte) 0x0;
    d[urlPos+11] = (byte) 0xaa;
    d[urlPos+12] = (byte) 0x0;
    d[urlPos+13] = (byte) 0x4b;
    d[urlPos+14] = (byte) 0xa9;
    d[urlPos+15] = (byte) 0x0b;

    // Number of characters in the url, including a zero trailing character
    IntegerHelper.getFourBytes((urlString.length() + 1)*2, d, urlPos+16);

    // Put the url into the data string
    StringHelper.getUnicodeBytes(urlString, d, urlPos+20);
    
    return d;    
  }

  /**
   * Gets the hyperlink stream specific to a URL link
   *
   * @param cd the data jxl.common.for all types of hyperlink
   * @return the raw data for a URL hyperlink
   */
  private byte[] getUNCData(byte[] cd)
  {
    String uncString = file.getPath();

    byte[] d = new byte[cd.length + uncString.length() * 2 + 2 + 4];
    System.arraycopy(cd, 0, d, 0, cd.length);

    int urlPos = cd.length;
    
    // The length of the unc string, including zero terminator
    int length = uncString.length() + 1;
    IntegerHelper.getFourBytes(length, d, urlPos);

    // Place the string into the stream
    StringHelper.getUnicodeBytes(uncString, d, urlPos + 4);

    return d;    
  }

  /**
   * Gets the hyperlink stream specific to a local file link
   *
   * @param cd the data jxl.common.for all types of hyperlink
   * @return the raw data for a URL hyperlink
   */
  private byte[] getFileData(byte[] cd)
  {
    // Build up the directory hierarchy in reverse order
    ArrayList path = new ArrayList();
    ArrayList shortFileName = new ArrayList();
    path.add(file.getName());
    shortFileName.add(getShortName(file.getName()));

    File parent = file.getParentFile();
    while (parent != null)
    {
      path.add(parent.getName());
      shortFileName.add(getShortName(parent.getName()));
      parent = parent.getParentFile();
    }

    // Deduce the up directory level count and remove the directory from
    // the path
    int upLevelCount = 0;
    int pos = path.size() - 1;
    boolean upDir = true;

    while (upDir)
    {
      String s = (String) path.get(pos);
      if (s.equals(".."))
      {
        upLevelCount++;
        path.remove(pos);
        shortFileName.remove(pos);
      }
      else
      {
        upDir = false;
      }

      pos--;
    }

    StringBuffer filePathSB = new StringBuffer();
    StringBuffer shortFilePathSB = new StringBuffer();

    if (file.getPath().charAt(1)==':')
    {
      char driveLetter = file.getPath().charAt(0);
      if (driveLetter != 'C' && driveLetter != 'c')
      {
        filePathSB.append(driveLetter);
        filePathSB.append(':');
        shortFilePathSB.append(driveLetter);
        shortFilePathSB.append(':');
      }
    }

    for (int i = path.size() - 1; i >= 0 ; i--)
    {
      filePathSB.append((String)path.get(i));
      shortFilePathSB.append((String)shortFileName.get(i));

      if (i != 0)
      {
        filePathSB.append("\\");
        shortFilePathSB.append("\\");
      }
    }


    String filePath = filePathSB.toString();
    String shortFilePath = shortFilePathSB.toString();

    int dataLength = cd.length + 
                     4 + (shortFilePath.length() + 1) + // short file name
                     16 + // inexplicable byte sequence
                     2 + // up directory level count
                     8 + (filePath.length() + 1) * 2 + // long file name
                     24; // inexplicable byte sequence


    if (contents != null)
    {
      dataLength += 4 + (contents.length() + 1) * 2;
    }

    // Copy across the jxl.common.data into the new array
    byte[] d = new byte[dataLength];

    System.arraycopy(cd, 0, d, 0, cd.length);
    
    int filePos = cd.length;
    
    // Add in the description text
    if (contents != null)
    {
      IntegerHelper.getFourBytes(contents.length() + 1, d, filePos);
      StringHelper.getUnicodeBytes(contents, d, filePos + 4);
      filePos += (contents.length() + 1) * 2 + 4;
    }

    int curPos = filePos;

    // Inexplicable byte sequence
    d[curPos]    = (byte) 0x03;
    d[curPos+1]  = (byte) 0x03;
    d[curPos+2]  = (byte) 0x0;
    d[curPos+3]  = (byte) 0x0;
    d[curPos+4]  = (byte) 0x0;
    d[curPos+5]  = (byte) 0x0;
    d[curPos+6]  = (byte) 0x0;
    d[curPos+7]  = (byte) 0x0;
    d[curPos+8]  = (byte) 0xc0;
    d[curPos+9]  = (byte) 0x0;
    d[curPos+10] = (byte) 0x0;
    d[curPos+11] = (byte) 0x0;
    d[curPos+12] = (byte) 0x0;
    d[curPos+13] = (byte) 0x0;
    d[curPos+14] = (byte) 0x0;
    d[curPos+15] = (byte) 0x46;

    curPos += 16;

    // The directory up level count
    IntegerHelper.getTwoBytes(upLevelCount, d, curPos);
    curPos += 2;

    // The number of bytes in the short file name, including zero terminator
    IntegerHelper.getFourBytes((shortFilePath.length() + 1), d, curPos);

    // The short file name
    StringHelper.getBytes(shortFilePath, d, curPos+4);

    curPos += 4 + (shortFilePath.length() + 1);

    // Inexplicable byte sequence
    d[curPos]   = (byte) 0xff;
    d[curPos+1] = (byte) 0xff;
    d[curPos+2] = (byte) 0xad;
    d[curPos+3] = (byte) 0xde;
    d[curPos+4] = (byte) 0x0;
    d[curPos+5] = (byte) 0x0;
    d[curPos+6] = (byte) 0x0;
    d[curPos+7] = (byte) 0x0;
    d[curPos+8] = (byte) 0x0;
    d[curPos+9] = (byte) 0x0;
    d[curPos+10] = (byte) 0x0;
    d[curPos+11] = (byte) 0x0;
    d[curPos+12] = (byte) 0x0;
    d[curPos+13] = (byte) 0x0;
    d[curPos+14] = (byte) 0x0;
    d[curPos+15] = (byte) 0x0;
    d[curPos+16] = (byte) 0x0;
    d[curPos+17] = (byte) 0x0;
    d[curPos+18] = (byte) 0x0;
    d[curPos+19] = (byte) 0x0;
    d[curPos+20] = (byte) 0x0;
    d[curPos+21] = (byte) 0x0;
    d[curPos+22] = (byte) 0x0;
    d[curPos+23] = (byte) 0x0;

    curPos += 24;

    // Size of the long file name data in bytes, including inexplicable data 
    // fields
    int size = 6 + filePath.length() * 2;
    IntegerHelper.getFourBytes(size, d, curPos);
    curPos += 4;

    // The number of bytes in the long file name
    // NOT including zero terminator
    IntegerHelper.getFourBytes((filePath.length()) * 2, d, curPos);
    curPos += 4;

    // Inexplicable bytes
    d[curPos] = (byte) 0x3;
    d[curPos+1] = (byte) 0x0;
    
    curPos += 2;

    // The long file name
    StringHelper.getUnicodeBytes(filePath, d, curPos);
    curPos += (filePath.length() + 1) * 2;
    

    /*
    curPos += 24;
    int nameLength = filePath.length() * 2;

    // Size of the file link 
    IntegerHelper.getFourBytes(nameLength+6, d, curPos);

    // Number of characters
    IntegerHelper.getFourBytes(nameLength, d, curPos+4);

    // Inexplicable byte sequence
    d[curPos+8] = 0x03;
    
    // The long file name
    StringHelper.getUnicodeBytes(filePath, d, curPos+10);
    */

    return d;    
  }

  /**
   * Gets the DOS short file name in 8.3 format of the name passed in
   * 
   * @param s the name
   * @return the dos short name
   */
  private String getShortName(String s)
  {
    int sep = s.indexOf('.');
    
    String prefix = null;
    String suffix = null;

    if (sep == -1)
    {
      prefix = s;
      suffix="";
    }
    else
    {
      prefix = s.substring(0,sep);
      suffix = s.substring(sep+1);
    }

    if (prefix.length() > 8)
    {
      prefix = prefix.substring(0, 6) + "~" + (prefix.length() - 8);
      prefix = prefix.substring(0, 8);
    }

    suffix = suffix.substring(0,Math.min(3, suffix.length()));

    if (suffix.length() > 0)
    {
      return prefix + '.' + suffix;
    }
    else
    {
      return prefix;
    }
  }

  /**
   * Gets the hyperlink stream specific to a location link
   *
   * @param cd the data jxl.common.for all types of hyperlink
   * @return the raw data for a URL hyperlink
   */
  private byte[] getLocationData(byte[] cd)
  {
    byte[] d = new byte[cd.length + 4 + (location.length() + 1)* 2];
    System.arraycopy(cd, 0, d, 0, cd.length);

    int locPos = cd.length;
    
    // The number of chars in the location string, plus a 0 terminator
    IntegerHelper.getFourBytes(location.length() + 1, d, locPos);
    
    // Get the location
    StringHelper.getUnicodeBytes(location, d, locPos+4);

    return d;    
  }

  
  /**
   * Initializes the range when this hyperlink is added to the sheet
   *
   * @param s the sheet containing this hyperlink
   */
  void initialize(WritableSheet s)
  {
    sheet = s;
    range = new SheetRangeImpl(s, 
                               firstColumn, firstRow,
                               lastColumn, lastRow);
  }

  /**
   * Called by the worksheet.  Gets the string contents to put into the cell
   * containing this hyperlink
   *
   * @return the string contents for the hyperlink cell
   */
  String getContents()
  {
    return contents;
  }

  /**
   * Sets the description
   *
   * @param desc the description
   */
  protected void setContents(String desc)
  {
    contents = desc;
    modified = true;
  }
}







