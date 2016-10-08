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

package jxl.biff;

import java.io.IOException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.format.Colour;
import jxl.format.RGB;
import jxl.write.biff.File;

/**
 * The list of XF records and formatting records for the workbook
 */
public class FormattingRecords
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(FormattingRecords.class);

  /**
   * A hash map of FormatRecords, for random access retrieval when reading
   * in a spreadsheet
   */
  private HashMap formats;

  /**
   * A list of formats, used when writing out a spreadsheet
   */
  private ArrayList formatsList;

  /**
   * The list of extended format records
   */
  private ArrayList xfRecords;

  /**
   * The next available index number for custom format records
   */
  private int nextCustomIndexNumber;

  /**
   * A handle to the available fonts
   */
  private Fonts fonts;

  /**
   * The colour palette
   */
  private PaletteRecord palette;

  /**
   * The start index number for custom format records
   */
  private static final int customFormatStartIndex = 0xa4;

  /**
   * The maximum number of format records.  This is some weird internal
   * Excel constraint
   */
  private static final int maxFormatRecordsIndex = 0x1b9;

  /**
   * The minimum number of XF records for a sheet.  The rationalization
   * processes commences immediately after this number
   */
  private static final int minXFRecords = 21;

  /**
   * Constructor
   *
   * @param f the container for the fonts
   */
  public FormattingRecords(Fonts f)
  {
    xfRecords = new ArrayList(10);
    formats = new HashMap(10);
    formatsList = new ArrayList(10);
    fonts = f;
    nextCustomIndexNumber = customFormatStartIndex;
  }

  /**
   * Adds an extended formatting record to the list.  If the XF record passed
   * in has not been initialized, its index is determined based on the
   * xfRecords list, and
   * this position is passed to the XF records initialize method
   *
   * @param xf the xf record to add
   * @exception NumFormatRecordsException
   */
  public final void addStyle(XFRecord xf)
    throws NumFormatRecordsException
  {
    if (!xf.isInitialized())
    {
      int pos = xfRecords.size();
      xf.initialize(pos, this, fonts);
      xfRecords.add(xf);
    }
    else
    {
      // The XF record has probably been read in.  If the index is greater
      // Than the size of the list, then it is not a preset format,
      // so add it
      if (xf.getXFIndex() >= xfRecords.size())
      {
        xfRecords.add(xf);
      }
    }
  }

  /**
   * Adds a cell format to the hash map, keyed on its index.  If the format
   * record is not initialized, then its index number is determined and its
   * initialize method called.  If the font is not a built in format, then it
   * is added to the list of formats for writing out
   *
   * @param fr the format record
   */
  public final void addFormat(DisplayFormat fr)
    throws NumFormatRecordsException
  {
    // Handle the case the where the index number in the read Excel
    // file exhibits some major weirdness
    if (fr.isInitialized() &&
        fr.getFormatIndex() >= maxFormatRecordsIndex)
    {
      logger.warn("Format index exceeds Excel maximum - assigning custom " +
                  "number");
      fr.initialize(nextCustomIndexNumber);
      nextCustomIndexNumber++;
    }

    // Initialize the format record with a custom index number
    if (!fr.isInitialized())
    {
      fr.initialize(nextCustomIndexNumber);
      nextCustomIndexNumber++;
    }

    if (nextCustomIndexNumber > maxFormatRecordsIndex)
    {
      nextCustomIndexNumber = maxFormatRecordsIndex;
      throw new NumFormatRecordsException();
    }

    if (fr.getFormatIndex() >= nextCustomIndexNumber)
    {
      nextCustomIndexNumber = fr.getFormatIndex() + 1;
    }

    if (!fr.isBuiltIn())
    {
      formatsList.add(fr);
      formats.put(new Integer(fr.getFormatIndex()), fr);
    }
  }

  /**
   * Sees if the extended formatting record at the specified position
   * represents a date.  First checks against the built in formats, and
   * then checks against the hash map of FormatRecords
   *
   * @param pos the xf format index
   * @return TRUE if this format index is formatted as a Date
   */
  public final boolean isDate(int pos)
  {
    XFRecord xfr = (XFRecord) xfRecords.get(pos);

    if (xfr.isDate())
    {
      return true;
    }

    FormatRecord fr = (FormatRecord)
      formats.get(new Integer(xfr.getFormatRecord()));

    return fr == null ? false : fr.isDate();
  }

  /**
   * Gets the DateFormat used to format the cell.
   *
   * @param pos the xf format index
   * @return the DateFormat object used to format the date in the original
   *     excel cell
   */
  public final DateFormat getDateFormat(int pos)
  {
    XFRecord xfr = (XFRecord) xfRecords.get(pos);

    if (xfr.isDate())
    {
      return xfr.getDateFormat();
    }

    FormatRecord fr = (FormatRecord)
      formats.get(new Integer(xfr.getFormatRecord()));

    if (fr == null)
    {
      return null;
    }

    return fr.isDate() ? fr.getDateFormat() : null;
  }

  /**
   * Gets the NumberFormat used to format the cell.
   *
   * @param pos the xf format index
   * @return the DateFormat object used to format the date in the original
   *     excel cell
   */
  public final NumberFormat getNumberFormat(int pos)
  {
    XFRecord xfr = (XFRecord) xfRecords.get(pos);
    
    if (xfr.isNumber())
    {
      return xfr.getNumberFormat();
    }

    FormatRecord fr = (FormatRecord)
      formats.get(new Integer(xfr.getFormatRecord()));

    if (fr == null)
    {
      return null;
    }

    return fr.isNumber() ? fr.getNumberFormat() : null;
  }

  /**
   * Gets the format record
   *
   * @param index the formatting record index to retrieve
   * @return the format record at the specified index
   */
  FormatRecord getFormatRecord(int index)
  {
    return (FormatRecord)
      formats.get(new Integer(index));
  }
  /**
   * Writes out all the format records and the XF records
   *
   * @param outputFile the file to write to
   * @exception IOException
   */
  public void write(File outputFile) throws IOException
  {
    // Write out all the formats
    Iterator i = formatsList.iterator();
    FormatRecord fr = null;
    while (i.hasNext())
    {
      fr = (FormatRecord) i.next();
      outputFile.write(fr);
    }

    // Write out the styles
    i = xfRecords.iterator();
    XFRecord xfr = null;
    while (i.hasNext())
    {
      xfr = (XFRecord) i.next();
      outputFile.write(xfr);
    }

    // Write out the style records
    BuiltInStyle style = new BuiltInStyle(0x10, 3);
    outputFile.write(style);

    style = new BuiltInStyle(0x11, 6);
    outputFile.write(style);

    style = new BuiltInStyle(0x12, 4);
    outputFile.write(style);

    style = new BuiltInStyle(0x13, 7);
    outputFile.write(style);

    style = new BuiltInStyle(0x0, 0);
    outputFile.write(style);

    style = new BuiltInStyle(0x14, 5);
    outputFile.write(style);
  }

  /**
   * Accessor for the fonts used by this workbook
   *
   * @return the fonts container
   */
  protected final Fonts getFonts()
  {
    return fonts;
  }

  /**
   * Gets the XFRecord for the specified index.  Used when copying individual
   * cells
   *
   * @param index the XF record to retrieve
   * @return the XF record at the specified index
   */
  public final XFRecord getXFRecord(int index)
  {
    return (XFRecord) xfRecords.get(index);
  }

  /**
   * Gets the number of formatting records on the list.  This is used by the
   * writable subclass because there is an upper limit on the amount of
   * format records that are allowed to be present in an excel sheet
   *
   * @return the number of format records present
   */
  protected final int getNumberOfFormatRecords()
  {
    return formatsList.size();
  }

  /**
   * Rationalizes all the fonts, removing duplicate entries
   *
   * @return the list of new font index number
   */
  public IndexMapping rationalizeFonts()
  {
    return fonts.rationalize();
  }

  /**
   * Rationalizes the cell formats.  Duplicate
   * formats are removed and the format indexed of the cells
   * adjusted accordingly
   *
   * @param fontMapping the font mapping index numbers
   * @param formatMapping the format mapping index numbers
   * @return the list of new font index number
   */
  public IndexMapping rationalize(IndexMapping fontMapping,
                                  IndexMapping formatMapping)
  {
    // Update the index codes for the XF records using the format
    // mapping and the font mapping
    // at the same time
    XFRecord xfr = null;
    for (Iterator it = xfRecords.iterator(); it.hasNext();)
    {
      xfr = (XFRecord) it.next();

      if (xfr.getFormatRecord() >= customFormatStartIndex)
      {
        xfr.setFormatIndex(formatMapping.getNewIndex(xfr.getFormatRecord()));
      }

      xfr.setFontIndex(fontMapping.getNewIndex(xfr.getFontIndex()));
    }

    ArrayList newrecords = new ArrayList(minXFRecords);
    IndexMapping mapping = new IndexMapping(xfRecords.size());
    int numremoved = 0;

    int numXFRecords = Math.min(minXFRecords, xfRecords.size());
    // Copy across the fundamental styles
    for (int i = 0; i < numXFRecords; i++)
    {
      newrecords.add(xfRecords.get(i));
      mapping.setMapping(i, i);
    }

    if (numXFRecords < minXFRecords)
    {
      logger.warn("There are less than the expected minimum number of " +
                   "XF records");
      return mapping;
    }

    // Iterate through the old list
    for (int i = minXFRecords; i < xfRecords.size(); i++)
    {
      XFRecord xf = (XFRecord) xfRecords.get(i);

      // Compare against formats already on the list
      boolean duplicate = false;
      for (Iterator it = newrecords.iterator();
           it.hasNext() && !duplicate;)
      {
        XFRecord xf2 = (XFRecord) it.next();
        if (xf2.equals(xf))
        {
          duplicate = true;
          mapping.setMapping(i, mapping.getNewIndex(xf2.getXFIndex()));
          numremoved++;
        }
      }

      // If this format is not a duplicate then add it to the new list
      if (!duplicate)
      {
        newrecords.add(xf);
        mapping.setMapping(i, i - numremoved);
      }
    }

    // It is sufficient to merely change the xf index field on all XFRecords
    // In this case, CellValues which refer to defunct format records
    // will nevertheless be written out with the correct index number
    for (Iterator i = xfRecords.iterator(); i.hasNext();)
    {
      XFRecord xf = (XFRecord) i.next();
      xf.rationalize(mapping);
    }

    // Set the new list
    xfRecords = newrecords;

    return mapping;
  }

  /**
   * Rationalizes the display formats.  Duplicate
   * formats are removed and the format indices of the cells
   * adjusted accordingly.  It is invoked immediately prior to writing
   * writing out the sheet
   * @return the index mapping between the old display formats and the
   * rationalized ones
   */
  public IndexMapping rationalizeDisplayFormats()
  {
    ArrayList newformats = new ArrayList();
    int numremoved = 0;
    IndexMapping mapping = new IndexMapping(nextCustomIndexNumber);

    // Iterate through the old list
    Iterator i = formatsList.iterator();
    DisplayFormat df = null;
    DisplayFormat df2 = null;
    boolean duplicate = false;
    while (i.hasNext())
    {
      df = (DisplayFormat) i.next();

      Assert.verify(!df.isBuiltIn());

      // Compare against formats already on the list
      Iterator i2 = newformats.iterator();
      duplicate = false;
      while (i2.hasNext() && !duplicate)
      {
        df2 = (DisplayFormat) i2.next();
        if (df2.equals(df))
        {
          duplicate = true;
          mapping.setMapping(df.getFormatIndex(),
                             mapping.getNewIndex(df2.getFormatIndex()));
          numremoved++;
        }
      }

      // If this format is not a duplicate then add it to the new list
      if (!duplicate)
      {
        newformats.add(df);
        int indexnum = df.getFormatIndex() - numremoved;
        if (indexnum > maxFormatRecordsIndex)
        {
          logger.warn("Too many number formats - using default format.");
          indexnum = 0; // the default number format index
        }
        mapping.setMapping(df.getFormatIndex(),
                           df.getFormatIndex() - numremoved);
      }
    }

    // Set the new list
    formatsList = newformats;

    // Update the index codes for the remaining formats
    i = formatsList.iterator();

    while (i.hasNext())
    {
      df = (DisplayFormat) i.next();
      df.initialize(mapping.getNewIndex(df.getFormatIndex()));
    }

    return mapping;
  }

  /**
   * Accessor for the colour palette
   *
   * @return the palette
   */
  public PaletteRecord getPalette()
  {
    return palette;
  }

  /**
   * Called from the WorkbookParser to set the colour palette
   *
   * @param pr the palette
   */
  public void setPalette(PaletteRecord pr)
  {
    palette = pr;
  }

  /**
   * Sets the RGB value for the specified colour for this workbook
   *
   * @param c the colour whose RGB value is to be overwritten
   * @param r the red portion to set (0-255)
   * @param g the green portion to set (0-255)
   * @param b the blue portion to set (0-255)
   */
  public void setColourRGB(Colour c, int r, int g, int b)
  {
    if (palette == null)
    {
      palette = new PaletteRecord();
    }
    palette.setColourRGB(c, r, g, b);
  }

  /**
   * Accessor for the RGB value for the specified colour
   *
   * @return the RGB for the specified colour
   */
  public RGB getColourRGB(Colour c)
  {
    if (palette == null)
    {
      return c.getDefaultRGB();
    }

    return palette.getColourRGB(c);
  }
}
