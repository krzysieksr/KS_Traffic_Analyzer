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

package jxl.read.biff;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;

import jxl.common.Logger;

import jxl.CellReferenceHelper;
import jxl.Hyperlink;
import jxl.Range;
import jxl.Sheet;
import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.SheetRangeImpl;
import jxl.biff.StringHelper;

/**
 * A number record.  This is stored as 8 bytes, as opposed to the
 * 4 byte RK record
 */
public class HyperlinkRecord extends RecordData implements Hyperlink
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
   * The range of cells which activate this hyperlink
   */
  private SheetRangeImpl range;

  /**
   * The type of this hyperlink
   */
  private LinkType linkType;

  /**
   * The excel type of hyperlink
   */
  private static class LinkType {};

  private static final LinkType urlLink      = new LinkType();
  private static final LinkType fileLink     = new LinkType();
  private static final LinkType workbookLink = new LinkType();
  private static final LinkType unknown      = new LinkType();

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param s the sheet
   * @param ws the workbook settings
   */
  HyperlinkRecord(Record t, Sheet s, WorkbookSettings ws)
  {
    super(t);

    linkType = unknown;

    byte[] data = getRecord().getData();

    // Build up the range of cells occupied by this hyperlink
    firstRow    = IntegerHelper.getInt(data[0], data[1]);
    lastRow     = IntegerHelper.getInt(data[2], data[3]);
    firstColumn = IntegerHelper.getInt(data[4], data[5]);
    lastColumn  = IntegerHelper.getInt(data[6], data[7]);
    range       = new SheetRangeImpl(s,
                                     firstColumn, firstRow,
                                     lastColumn, lastRow);

    int options = IntegerHelper.getInt(data[28], data[29], data[30], data[31]);

    boolean description = (options & 0x14) != 0;
    int startpos = 32;
    int descbytes = 0;
    if (description)
    {
      int descchars = IntegerHelper.getInt
        (data[startpos], data[startpos + 1],
         data[startpos + 2], data[startpos + 3]);
      descbytes = descchars * 2 + 4;
    }

    startpos += descbytes;

    boolean targetFrame = (options & 0x80) != 0;
    int targetbytes = 0;
    if (targetFrame)
    {
      int targetchars = IntegerHelper.getInt
        (data[startpos], data[startpos + 1],
         data[startpos + 2], data[startpos + 3]);
      targetbytes = targetchars * 2 + 4;
    }

    startpos += targetbytes;

    // Try and determine the type
    if ((options & 0x3) == 0x03)
    {
      linkType = urlLink;

      // check the guid monicker
      if (data[startpos] == 0x03)
      {
        linkType = fileLink;
      }
    }
    else if ((options & 0x01) != 0)
    {
      linkType = fileLink;
      // check the guid monicker
      if (data[startpos] == (byte) 0xe0)
      {
        linkType = urlLink;
      }
    }
    else if ((options & 0x08) != 0)
    {
      linkType = workbookLink;
    }

    // Try and determine the type
    if (linkType == urlLink)
    {
      String urlString = null;
      try
      {
        startpos += 16;

        // Get the url, ignoring the 0 char at the end
        int bytes = IntegerHelper.getInt(data[startpos],
                                         data[startpos + 1],
                                         data[startpos + 2],
                                         data[startpos + 3]);

        urlString = StringHelper.getUnicodeString(data, bytes / 2 - 1,
                                                  startpos + 4);
        url = new URL(urlString);
      }
      catch (MalformedURLException e)
      {
        logger.warn("URL " + urlString + " is malformed.  Trying a file");
        try
        {
          linkType = fileLink;
          file = new File(urlString);
        }
        catch (Exception e3)
        {
          logger.warn("Cannot set to file.  Setting a default URL");

          // Set a default URL
          try
          {
            linkType = urlLink;
            url = new URL("http://www.andykhan.com/jexcelapi/index.html");
          }
          catch (MalformedURLException e2)
          {
            // fail silently
          }
        }
      }
      catch (Throwable e)
      {
        StringBuffer sb1 = new StringBuffer();
        StringBuffer sb2 = new StringBuffer();
        CellReferenceHelper.getCellReference(firstColumn, firstRow, sb1);
        CellReferenceHelper.getCellReference(lastColumn, lastRow, sb2);
        sb1.insert(0, "Exception when parsing URL ");
        sb1.append('\"').append(sb2.toString()).append("\".  Using default.");
        logger.warn(sb1, e);

        // Set a default URL
        try
        {
          url = new URL("http://www.andykhan.com/jexcelapi/index.html");
        }
        catch (MalformedURLException e2)
        {
          // fail silently
        }
      }
    }
    else if (linkType == fileLink)
    {
      try
      {
        startpos += 16;

        // Get the name of the local file, ignoring the zero character at the
        // end
        int upLevelCount = IntegerHelper.getInt(data[startpos],
                                                data[startpos + 1]);
        int chars = IntegerHelper.getInt(data[startpos + 2],
                                         data[startpos + 3],
                                         data[startpos + 4],
                                         data[startpos + 5]);
        String fileName = StringHelper.getString(data, chars - 1,
                                                 startpos + 6, ws);

        StringBuffer sb = new StringBuffer();

        for (int i = 0; i < upLevelCount; i++)
        {
          sb.append("..\\");
        }

        sb.append(fileName);

        file = new File(sb.toString());
      }
      catch (Throwable e)
      {
        logger.warn("Exception when parsing file " + 
                    e.getClass().getName() + ".");
        file = new File(".");
      }
    }
    else if (linkType == workbookLink)
    {
      int chars = IntegerHelper.getInt(data[32], data[33], data[34], data[35]);
      location  = StringHelper.getUnicodeString(data, chars - 1, 36);
    }
    else
    {
      // give up
      logger.warn("Cannot determine link type");
      return;
    }
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
   * Exposes the base class method.  This is used when copying hyperlinks
   *
   * @return the Record data
   */
  public Record getRecord()
  {
    return super.getRecord();
  }

  /**
   * Gets the range of cells which activate this hyperlink
   * The get sheet index methods will all return -1, because the
   * cells will all be present on the same sheet
   *
   * @return the range of cells which activate the hyperlink
   */
  public Range getRange()
  {
    return range;
  }

  /**
   * Gets the location referenced by this hyperlink
   *
   * @return the location
   */
  public String getLocation()
  {
    return location;
  }
}







