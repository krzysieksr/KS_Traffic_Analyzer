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

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.Format;
import jxl.format.Orientation;
import jxl.format.Pattern;
import jxl.format.VerticalAlignment;
import jxl.read.biff.Record;

/**
 * Holds an extended formatting record
 */
public class XFRecord extends WritableRecordData implements CellFormat
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(XFRecord.class);

  /**
   * The index to the format record
   */
  public int formatIndex;

  /**
   * The index of the parent format
   */
  private int parentFormat;

  /**
   * The format type
   */
  private XFType xfFormatType;

  /**
   * Indicates whether this is a date formatting record
   */
  private boolean date;

  /**
   * Indicates whether this is a number formatting record
   */
  private boolean number;

  /**
   * The date format for this record.  Deduced when the record is
   * read in from a spreadsheet
   */
  private DateFormat dateFormat;

  /**
   * The number format for this record.  Deduced when the record is read in
   * from a spreadsheet
   */
  private NumberFormat numberFormat;

  /**
   * The used attribute.  Needs to be preserved in order to get accurate
   * rationalization
   */
  private byte usedAttributes;
  /**
   * The index to the font record used by this XF record
   */
  private int fontIndex;
  /**
   * Flag to indicate whether this XF record represents a locked cell
   */
  private boolean locked;
  /**
   * Flag to indicate whether this XF record is hidden
   */
  private boolean hidden;
  /**
   * The alignment for this cell (left, right, centre)
   */
  private Alignment align;
  /**
   * The vertical alignment for the cell (top, bottom, centre)
   */
  private VerticalAlignment valign;
  /**
   * The orientation of the cell
   */
  private Orientation orientation;
  /**
   * Flag to indicates whether the data (normally text) in the cell will be
   * wrapped around to fit in the cell width
   */
  private boolean wrap;

  /**
   * Indentation of the cell text
   */
  private int indentation;

  /**
   * Flag to indicate that this format is shrink to fit
   */
  private boolean shrinkToFit;

  /**
   * The border indicator for the left of this cell
   */
  private BorderLineStyle leftBorder;
  /**
   * The border indicator for the right of the cell
   */
  private BorderLineStyle rightBorder;
  /**
   * The border indicator for the top of the cell
   */
  private BorderLineStyle topBorder;
  /**
   * The border indicator for the bottom of the cell
   */
  private BorderLineStyle bottomBorder;

  /**
   * The border colour for the left of the cell
   */
  private Colour leftBorderColour;
  /**
   * The border colour for the right of the cell
   */
  private Colour rightBorderColour;
  /**
   * The border colour for the top of the cell
   */
  private Colour topBorderColour;
  /**
   * The border colour for the bottom of the cell
   */
  private Colour bottomBorderColour;

  /**
   * The background colour
   */
  private Colour backgroundColour;
  /**
   * The background pattern
   */
  private Pattern pattern;
  /**
   * The options mask which is used to store the processed cell options (such
   * as alignment, borders etc)
   */
  private int options;
  /**
   * The index of this XF record within the workbook
   */
  private int xfIndex;
  /**
   * The font object for this XF record
   */
  private FontRecord font;
  /**
   * The format object for this XF record.  This is used when creating
   * a writable record
   */
  private DisplayFormat format;
  /**
   * Flag to indicate whether this XF record has been initialized
   */
  private boolean initialized;

  /**
   * Indicates whether this cell was constructed by an API or read
   * from an existing Excel file
   */
  private boolean read;

  /**
   * The excel format for this record. This is used to display the actual
   * excel format string back to the user (eg. when generating certain
   * types of XML document) as opposed to the java equivalent
   */
  private Format excelFormat;

  /**
   * Flag to indicate whether the format information has been initialized.
   * This is false if the xf record has been read in, but true if it
   * has been written
   */
  private boolean formatInfoInitialized;

  /**
   * Flag to indicate whether this cell was copied.  If it was copied, then
   * it can be set to uninitialized, allowing us to change certain format
   * information
   */
  private boolean copied;

  /**
   * A handle to the formatting records.  The purpose of this is
   * to read the formatting information back, for the purposes of
   * output eg. to some form of XML
   */
  private FormattingRecords formattingRecords;

  /** 
   * Constants for the used attributes
   */
  private static final int USE_FONT = 0x4;
  private static final int USE_FORMAT = 0x8;
  private static final int USE_ALIGNMENT = 0x10;
  private static final int USE_BORDER = 0x20;
  private static final int USE_BACKGROUND = 0x40;
  private static final int USE_PROTECTION = 0x80;
  private static final int USE_DEFAULT_VALUE=0xf8;

  /**
   * The list of built in date format values
   */
  private static final int[] dateFormats = new int[]
    {0xe,
     0xf,
     0x10,
     0x11,
     0x12,
     0x13,
     0x14,
     0x15,
     0x16,
     0x2d,
     0x2e,
     0x2f};

  /**
   * The list of java date format equivalents
   */
  private static final DateFormat[] javaDateFormats = new DateFormat[]
    {SimpleDateFormat.getDateInstance(DateFormat.SHORT),
     SimpleDateFormat.getDateInstance(DateFormat.MEDIUM),
     new SimpleDateFormat("d-MMM"),
     new SimpleDateFormat("MMM-yy"),
     new SimpleDateFormat("h:mm a"),
     new SimpleDateFormat("h:mm:ss a"),
     new SimpleDateFormat("H:mm"),
     new SimpleDateFormat("H:mm:ss"),
     new SimpleDateFormat("M/d/yy H:mm"),
     new SimpleDateFormat("mm:ss"),
     new SimpleDateFormat("H:mm:ss"),
     new SimpleDateFormat("mm:ss.S")};

  /**
   * The list of built in number format values
   */
  private static int[] numberFormats = new int[]
    {0x1,
     0x2,
     0x3,
     0x4,
     0x5,
     0x6,
     0x7,
     0x8,
     0x9,
     0xa,
     0xb,
     0x25,
     0x26,
     0x27,
     0x28,
     0x29,
     0x2a,
     0x2b,
     0x2c,
     0x30};

  /**
   *  The list of java number format equivalents
   */
  private static NumberFormat[] javaNumberFormats = new NumberFormat[]
    {new DecimalFormat("0"),
     new DecimalFormat("0.00"),
     new DecimalFormat("#,##0"),
     new DecimalFormat("#,##0.00"),
     new DecimalFormat("$#,##0;($#,##0)"),
     new DecimalFormat("$#,##0;($#,##0)"),
     new DecimalFormat("$#,##0.00;($#,##0.00)"),
     new DecimalFormat("$#,##0.00;($#,##0.00)"),
     new DecimalFormat("0%"),
     new DecimalFormat("0.00%"),
     new DecimalFormat("0.00E00"),
     new DecimalFormat("#,##0;(#,##0)"),
     new DecimalFormat("#,##0;(#,##0)"),
     new DecimalFormat("#,##0.00;(#,##0.00)"),
     new DecimalFormat("#,##0.00;(#,##0.00)"),
     new DecimalFormat("#,##0;(#,##0)"),
     new DecimalFormat("$#,##0;($#,##0)"),
     new DecimalFormat("#,##0.00;(#,##0.00)"),
     new DecimalFormat("$#,##0.00;($#,##0.00)"),
     new DecimalFormat("##0.0E0")};

  // Type to distinguish between biff7 and biff8
  private static class BiffType {};

  public static final BiffType biff8 = new BiffType();
  public static final BiffType biff7 = new BiffType();

  /**
   * The biff type
   */
  private BiffType biffType;

  // Type to distinguish between cell and style records
  private static class XFType
  {
  }
  protected static final XFType cell = new XFType();
  protected static final XFType style = new XFType();

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param bt the biff type
   */
  public XFRecord(Record t, WorkbookSettings ws, BiffType bt)
  {
    super(t);

    biffType = bt;

    byte[] data = getRecord().getData();

    fontIndex = IntegerHelper.getInt(data[0], data[1]);
    formatIndex = IntegerHelper.getInt(data[2], data[3]);
    date = false;
    number = false;


    // Compare against the date formats
    for (int i = 0; i < dateFormats.length && date == false; i++)
    {
      if (formatIndex == dateFormats[i])
      {
        date = true;
        dateFormat = javaDateFormats[i];
      }
    }

    // Compare against the number formats
    for (int i = 0; i < numberFormats.length && number == false; i++)
    {
      if (formatIndex == numberFormats[i])
      {
        number = true;
        DecimalFormat df = (DecimalFormat) javaNumberFormats[i].clone();
        DecimalFormatSymbols symbols = 
          new DecimalFormatSymbols(ws.getLocale());
        df.setDecimalFormatSymbols(symbols);
        numberFormat = df;
        //numberFormat = javaNumberFormats[i];
      }
    }

    // Initialize the parent format and the type
    int cellAttributes = IntegerHelper.getInt(data[4], data[5]);
    parentFormat = (cellAttributes & 0xfff0) >> 4;

    int formatType = cellAttributes & 0x4;
    xfFormatType = formatType == 0 ? cell : style;
    locked = ((cellAttributes & 0x1) != 0);
    hidden = ((cellAttributes & 0x2) != 0);

    if (xfFormatType == cell &&
        (parentFormat & 0xfff) == 0xfff)
    {
      // Something is screwy with the parent format - set to zero
      parentFormat = 0;
      logger.warn("Invalid parent format found - ignoring");
    }

    initialized = false;
    read = true;
    formatInfoInitialized = false;
    copied = false;
  }

  /**
   * A constructor used when creating a writable record
   *
   * @param fnt the font
   * @param form the format
   */
  public XFRecord(FontRecord fnt, DisplayFormat form)
  {
    super(Type.XF);

    initialized        = false;
    locked             = true;
    hidden             = false;
    align              = Alignment.GENERAL;
    valign             = VerticalAlignment.BOTTOM;
    orientation        = Orientation.HORIZONTAL;
    wrap               = false;
    leftBorder         = BorderLineStyle.NONE;
    rightBorder        = BorderLineStyle.NONE;
    topBorder          = BorderLineStyle.NONE;
    bottomBorder       = BorderLineStyle.NONE;
    leftBorderColour   = Colour.AUTOMATIC;
    rightBorderColour  = Colour.AUTOMATIC;
    topBorderColour    = Colour.AUTOMATIC;
    bottomBorderColour = Colour.AUTOMATIC;
    pattern            = Pattern.NONE;
    backgroundColour   = Colour.DEFAULT_BACKGROUND;
    indentation        = 0;
    shrinkToFit        = false;
    usedAttributes     = (byte) (USE_FONT | USE_FORMAT | 
                         USE_BACKGROUND | USE_ALIGNMENT | USE_BORDER);

    // This will be set by the initialize method and the subclass respectively
    parentFormat = 0;
    xfFormatType = null;

    font     = fnt;
    format   = form;
    biffType = biff8;
    read     = false;
    copied   = false;
    formatInfoInitialized = true;

    Assert.verify(font != null);
    Assert.verify(format != null);
  }

  /**
   * Copy constructor.  Used for copying writable formats, typically
   * when duplicating formats to handle merged cells
   *
   * @param fmt XFRecord
   */
  protected XFRecord(XFRecord fmt)
  {
    super(Type.XF);

    initialized  = false;
    locked       = fmt.locked;
    hidden       = fmt.hidden;
    align        = fmt.align;
    valign       = fmt.valign;
    orientation  = fmt.orientation;
    wrap         = fmt.wrap;
    leftBorder   = fmt.leftBorder;
    rightBorder  = fmt.rightBorder;
    topBorder    = fmt.topBorder;
    bottomBorder = fmt.bottomBorder;
    leftBorderColour   = fmt.leftBorderColour;
    rightBorderColour  = fmt.rightBorderColour;
    topBorderColour    = fmt.topBorderColour;
    bottomBorderColour = fmt.bottomBorderColour;
    pattern            = fmt.pattern;
    xfFormatType       = fmt.xfFormatType;
    indentation        = fmt.indentation;
    shrinkToFit        = fmt.shrinkToFit;
    parentFormat       = fmt.parentFormat;
    backgroundColour   = fmt.backgroundColour;

    // Shallow copy is sufficient for these purposes
    font = fmt.font;
    format = fmt.format;

    fontIndex   = fmt.fontIndex;
    formatIndex = fmt.formatIndex;

    formatInfoInitialized = fmt.formatInfoInitialized;

    biffType = biff8;
    read = false;
    copied = true;
  }

  /**
   * A public copy constructor which can be used for copy formats between
   * different sheets.  Unlike the the other copy constructor, this
   * version does a deep copy
   *
   * @param cellFormat the format to copy
   */
  protected XFRecord(CellFormat cellFormat)
  {
    super(Type.XF);

    Assert.verify(cellFormat != null);
    Assert.verify(cellFormat instanceof XFRecord);
    XFRecord fmt = (XFRecord) cellFormat;

    if (!fmt.formatInfoInitialized)
    {
      fmt.initializeFormatInformation();
    }

    locked             = fmt.locked;
    hidden             = fmt.hidden;
    align              = fmt.align;
    valign             = fmt.valign;
    orientation        = fmt.orientation;
    wrap               = fmt.wrap;
    leftBorder         = fmt.leftBorder;
    rightBorder        = fmt.rightBorder;
    topBorder          = fmt.topBorder;
    bottomBorder       = fmt.bottomBorder;
    leftBorderColour   = fmt.leftBorderColour;
    rightBorderColour  = fmt.rightBorderColour;
    topBorderColour    = fmt.topBorderColour;
    bottomBorderColour = fmt.bottomBorderColour;
    pattern            = fmt.pattern;
    xfFormatType       = fmt.xfFormatType;
    parentFormat       = fmt.parentFormat;
    indentation        = fmt.indentation;
    shrinkToFit        = fmt.shrinkToFit;
    backgroundColour   = fmt.backgroundColour;

    // Deep copy of the font
    font = new FontRecord(fmt.getFont());

    // Copy the format
    if (fmt.getFormat() == null)
    {
      // format is writable
      if (fmt.format.isBuiltIn())
      {
        format = fmt.format;
      }
      else
      {
        // Format is not built in, so do a deep copy
        format = new FormatRecord((FormatRecord) fmt.format);
      }
    }
    else if (fmt.getFormat() instanceof BuiltInFormat)
    {
      // read excel format is built in
      excelFormat = (BuiltInFormat) fmt.excelFormat;
      format = (BuiltInFormat) fmt.excelFormat;
    }
    else
    {
      // read excel format is user defined
      Assert.verify(fmt.formatInfoInitialized);

      // in this case FormattingRecords should initialize the excelFormat
      // field with an instance of FormatRecord
      Assert.verify(fmt.excelFormat instanceof FormatRecord);

      // Format is not built in, so do a deep copy
      FormatRecord fr = new FormatRecord((FormatRecord) fmt.excelFormat);

      // Set both format fields to be the same object, since
      // FormatRecord implements all the necessary interfaces
      excelFormat = fr;
      format = fr;
    }

    biffType = biff8;

    // The format info should be all OK by virtue of the deep copy
    formatInfoInitialized = true;

    // This format was not read in
    read = false;

    // Treat this as a new cell record, so set the copied flag to false
    copied = false;

    // The font or format indexes need to be set, so set initialized to false
    initialized = false;
  }

  /**
   * Gets the java date format for this format record
   *
   * @return returns the date format
   */
  public DateFormat getDateFormat()
  {
    return dateFormat;
  }

  /**
   * Gets the java number format for this format record
   *
   * @return returns the number format
   */
  public NumberFormat getNumberFormat()
  {
    return numberFormat;
  }

  /**
   * Gets the lookup number of the format record
   *
   * @return returns the lookup number of the format record
   */
  public int getFormatRecord()
  {
    return formatIndex;
  }

  /**
   * Sees if this format is a date format
   *
   * @return TRUE if this refers to a built in date format
   */
  public boolean isDate()
  {
    return date;
  }

  /**
   * Sees if this format is a number format
   *
   * @return TRUE if this refers to a built in date format
   */
  public boolean isNumber()
  {
    return number;
  }

  /**
   * Converts the various fields into binary data.  If this object has
   * been read from an Excel file rather than being requested by a user (ie.
   * if the read flag is TRUE) then
   * no processing takes place and the raw data is simply returned.
   *
   * @return the raw data for writing
   */
  public byte[] getData()
  {
    // Format rationalization process means that we always want to
    // regenerate the format info - even if the spreadsheet was
    // read in
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    byte[] data = new byte[20];

    IntegerHelper.getTwoBytes(fontIndex, data, 0);
    IntegerHelper.getTwoBytes(formatIndex, data, 2);

    // Do the cell attributes
    int cellAttributes = 0;

    if (getLocked())
    {
      cellAttributes |= 0x01;
    }

    if (getHidden())
    {
      cellAttributes |= 0x02;
    }

    if (xfFormatType == style)
    {
      cellAttributes |= 0x04;
      parentFormat = 0xffff;
    }

    cellAttributes |= (parentFormat << 4);

    IntegerHelper.getTwoBytes(cellAttributes, data, 4);

    int alignMask = align.getValue();

    if (wrap)
    {
      alignMask |= 0x08;
    }

    alignMask |= (valign.getValue() << 4);

    alignMask |= (orientation.getValue() << 8);

    IntegerHelper.getTwoBytes(alignMask, data, 6);

    data[9] = (byte) 0x10;

    // Set the borders
    int borderMask = leftBorder.getValue();
    borderMask |= (rightBorder.getValue() << 4);
    borderMask |= (topBorder.getValue() << 8);
    borderMask |= (bottomBorder.getValue() << 12);

    IntegerHelper.getTwoBytes(borderMask, data, 10);

    // Set the border palette information if border mask is non zero
    // Hard code the colours to be black
    if (borderMask != 0)
    {
    	byte lc = (byte)leftBorderColour.getValue();
    	byte rc = (byte)rightBorderColour.getValue();
    	byte tc = (byte)topBorderColour.getValue();
    	byte bc = (byte)bottomBorderColour.getValue();

      int sideColourMask = (lc & 0x7f) | ((rc & 0x7f) << 7);
      int topColourMask  = (tc & 0x7f) | ((bc & 0x7f) << 7);

      IntegerHelper.getTwoBytes(sideColourMask, data, 12);
      IntegerHelper.getTwoBytes(topColourMask, data, 14);
    }

    // Set the background pattern
    int patternVal = pattern.getValue() << 10;
    IntegerHelper.getTwoBytes(patternVal, data, 16);

    // Set the colour palette
    int colourPaletteMask = backgroundColour.getValue();
    colourPaletteMask |= (0x40 << 7);
    IntegerHelper.getTwoBytes(colourPaletteMask, data, 18);

    // Set the cell options
    options |= indentation & 0x0f;

    if (shrinkToFit)
    {
      options |= 0x10;
    }
    else
    {
      options &= 0xef;
    }

    data[8] = (byte) options;

    if (biffType == biff8)
    {
      data[9] = (byte) usedAttributes;
    }

    return data;
  }

  /**
   * Accessor for the locked flag
   *
   * @return TRUE if this XF record locks cells, FALSE otherwise
   */
  protected final boolean getLocked()
  {
    return locked;
  }

  /**
   * Accessor for the hidden flag
   *
   * @return TRUE if this XF record hides the cell, FALSE otherwise
   */
  protected final boolean getHidden()
  {
    return hidden;
  }

  /**
   * Sets whether or not this XF record locks the cell
   *
   * @param l the locked flag
   */
  protected final void setXFLocked(boolean l)
  {
    locked = l;
    usedAttributes |= USE_PROTECTION;
  }

  /**
   * Sets the cell options
   *
   * @param opt the cell options
   */
  protected final void setXFCellOptions(int opt)
  {
    options |= opt;
  }

  /**
   * Sets the horizontal alignment for the data in this cell.
   * This method should only be called from its writable subclass
   * CellXFRecord
   *
   * @param a the alignment
   */
  protected void setXFAlignment(Alignment a)
  {
    Assert.verify(!initialized);
    align = a;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Sets the indentation
   *
   * @param i the indentation
   */
  protected void setXFIndentation(int i)
  {
    Assert.verify(!initialized);
    indentation = i;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Sets the shrink to fit flag
   *
   * @param s the shrink to fit flag
   */
  protected void setXFShrinkToFit(boolean s)
  {
    Assert.verify(!initialized);
    shrinkToFit = s;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Gets the horizontal cell alignment
   *
   * @return the alignment
   */
  public Alignment getAlignment()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return align;
  }

  /**
   * Gets the indentation
   *
   * @return the indentation
   */
  public int getIndentation() 
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return indentation;
  }

  /**
   * Gets the shrink to fit flag
   *
   * @return TRUE if this format is shrink to fit, FALSE otherise
   */
  public boolean isShrinkToFit()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return shrinkToFit;
  }

  /**
   * Accessor for whether a particular cell is locked
   *
   * @return TRUE if this cell is locked, FALSE otherwise
   */
  public boolean isLocked()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return locked;
  }


  /**
   * Gets the vertical cell alignment
   *
   * @return the alignment
   */
  public VerticalAlignment getVerticalAlignment()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return valign;
  }

  /**
   * Gets the orientation
   *
   * @return the orientation
   */
  public Orientation getOrientation()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return orientation;
  }

  /**
   * Sets the horizontal alignment for the data in this cell.
   * This method should only be called from its writable subclass
   * CellXFRecord
   *
   * @param c the background colour
   * @param p the background pattern
   */
  protected void setXFBackground(Colour c, Pattern p)
  {
    Assert.verify(!initialized);
    backgroundColour = c;
    pattern = p;
    usedAttributes |= USE_BACKGROUND;
  }

  /**
   * Gets the background colour used by this cell
   *
   * @return the foreground colour
   */
  public Colour getBackgroundColour()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return backgroundColour;
  }

  /**
   * Gets the pattern used by this cell format
   *
   * @return the background pattern
   */
  public Pattern getPattern()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return pattern;
  }

  /**
   * Sets the vertical alignment for the data in this cell
   * This method should only be called from its writable subclass
   * CellXFRecord

   *
   * @param va the vertical alignment
   */
  protected void setXFVerticalAlignment(VerticalAlignment va)
  {
    Assert.verify(!initialized);
    valign = va;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Sets the vertical alignment for the data in this cell
   * This method should only be called from its writable subclass
   * CellXFRecord

   *
   * @param o the orientation
   */
  protected void setXFOrientation(Orientation o)
  {
    Assert.verify(!initialized);
    orientation = o;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Sets whether the data in this cell is wrapped
   * This method should only be called from its writable subclass
   * CellXFRecord
   *
   * @param w the wrap flag
   */
  protected void setXFWrap(boolean w)
  {
    Assert.verify(!initialized);
    wrap = w;
    usedAttributes |= USE_ALIGNMENT;
  }

  /**
   * Gets whether or not the contents of this cell are wrapped
   *
   * @return TRUE if this cell's contents are wrapped, FALSE otherwise
   */
  public boolean getWrap()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    return wrap;
  }

  /**
   * Sets the border for this cell
   * This method should only be called from its writable subclass
   * CellXFRecord
   *
   * @param b the border
   * @param ls the border line style
   */
  protected void setXFBorder(Border b, BorderLineStyle ls, Colour c)
  {
    Assert.verify(!initialized);
    
    if (c == Colour.BLACK || c == Colour.UNKNOWN) 
    {
      c = Colour.PALETTE_BLACK;
    }

    if (b == Border.LEFT)
    {
      leftBorder = ls;
      leftBorderColour = c;
    }
    else if (b == Border.RIGHT)
    {
      rightBorder = ls;
      rightBorderColour = c;
    }
    else if (b == Border.TOP)
    {
      topBorder = ls;
      topBorderColour = c;
    }
    else if (b == Border.BOTTOM)
    {
      bottomBorder = ls;
      bottomBorderColour = c;
    }

    usedAttributes |= USE_BORDER;

    return;
  }


  /**
   * Gets the line style for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public BorderLineStyle getBorder(Border border)
  {
    return getBorderLine(border);
  }

  /**
   * Gets the line style for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public BorderLineStyle getBorderLine(Border border)
  {
    // Don't bother with the short cut records
    if (border == Border.NONE ||
        border == Border.ALL)
    {
      return BorderLineStyle.NONE;
    }

    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    if (border == Border.LEFT)
    {
      return leftBorder;
    }
    else if (border == Border.RIGHT)
    {
      return rightBorder;
    }
    else if (border == Border.TOP)
    {
      return topBorder;
    }
    else if (border == Border.BOTTOM)
    {
      return bottomBorder;
    }

    return BorderLineStyle.NONE;
  }

  /**
   * Gets the line style for the given cell border
   * If a border type of ALL or NONE is specified, then a line style of
   * NONE is returned
   *
   * @param border the cell border we are interested in
   * @return the line style of the specified border
   */
  public Colour getBorderColour(Border border)
  {
    // Don't bother with the short cut records
    if (border == Border.NONE ||
        border == Border.ALL)
    {
      return Colour.PALETTE_BLACK;
    }

    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    if (border == Border.LEFT)
    {
      return leftBorderColour;
    }
    else if (border == Border.RIGHT)
    {
      return rightBorderColour;
    }
    else if (border == Border.TOP)
    {
      return topBorderColour;
    }
    else if (border == Border.BOTTOM)
    {
      return bottomBorderColour;
    }

    return Colour.BLACK;  	
  }


  /**
   * Determines if this cell format has any borders at all.  Used to
   * set the new borders when merging a group of cells
   *
   * @return TRUE if this cell has any borders, FALSE otherwise
   */
  public final boolean hasBorders()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    if (leftBorder   == BorderLineStyle.NONE &&
        rightBorder  == BorderLineStyle.NONE &&
        topBorder    == BorderLineStyle.NONE &&
        bottomBorder == BorderLineStyle.NONE)
    {
      return false;
    }

    return true;
  }

  /**
   * If this cell has not been read in from an existing Excel sheet,
   * then initializes this record with the XF index passed in. Calls
   * initialized on the font and format record
   *
   * @param pos the xf index to initialize this record with
   * @param fr the containing formatting records
   * @param fonts the container for the fonts
   * @exception NumFormatRecordsException
   */
  public final void initialize(int pos, FormattingRecords fr, Fonts fonts)
    throws NumFormatRecordsException
  {
    xfIndex = pos;
    formattingRecords = fr;

    // If this file has been read in or copied,
    // the font and format indexes will
    // already be initialized, so just set the initialized flag and
    // return
    if (read || copied)
    {
      initialized = true;
      return;
    }

    if (!font.isInitialized())
    {
      fonts.addFont(font);
    }

    if (!format.isInitialized())
    {
      fr.addFormat(format);
    }

    fontIndex = font.getFontIndex();
    formatIndex = format.getFormatIndex();

    initialized = true;
  }

  /**
   * Resets the initialize flag.  This is called by the constructor of
   * WritableWorkbookImpl to reset the statically declared fonts
   */
  public final void uninitialize()
  {
    // As the default formats are cloned internally, the initialized
    // flag should never be anything other than false
    if (initialized == true)
    {
      logger.warn("A default format has been initialized");
    }
    initialized = false;
  }

  /**
   * Sets the XF index.  Called when rationalizing the XF records
   * immediately prior to writing
   *
   * @param xfi the new xf index
   */
  final void setXFIndex(int xfi)
  {
    xfIndex = xfi;
  }

  /**
   * Accessor for the XF index
   *
   * @return the XF index for this cell
   */
  public final int getXFIndex()
  {
    return xfIndex;
  }

  /**
   * Accessor to see if this format is initialized
   *
   * @return TRUE if this format is initialized, FALSE otherwise
   */
  public final boolean isInitialized()
  {
    return initialized;
  }

  /**
   * Accessor to see if this format was read in.  Used when checking merged
   * cells
   *
   * @return TRUE if this XF record was read in, FALSE if it was generated by
   *         the user API
   */
  public final boolean isRead()
  {
    return read;
  }

  /**
   * Gets the format used by this format
   *
   * @return the format
   */
  public Format getFormat()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }
    return excelFormat;
  }

  /**
   * Gets the font used by this format
   *
   * @return the font
   */
  public Font getFont()
  {
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }
    return font;
  }

  /**
   * Initializes the internal format information from the data read in
   */
  private void initializeFormatInformation()
  {
    // Initialize the cell format string
    if (formatIndex < BuiltInFormat.builtIns.length &&
        BuiltInFormat.builtIns[formatIndex] != null)
    {
      excelFormat = BuiltInFormat.builtIns[formatIndex];
    }
    else
    {
      excelFormat = formattingRecords.getFormatRecord(formatIndex);
    }

    // Initialize the font
    font = formattingRecords.getFonts().getFont(fontIndex);

    // Initialize the cell format data from the binary record
    byte[] data = getRecord().getData();

    // Get the parent record
    int cellAttributes = IntegerHelper.getInt(data[4], data[5]);
    parentFormat = (cellAttributes & 0xfff0) >> 4;
    int formatType = cellAttributes & 0x4;
    xfFormatType = formatType == 0 ? cell : style;
    locked = ((cellAttributes & 0x1) != 0);
    hidden = ((cellAttributes & 0x2) != 0);

    if (xfFormatType == cell &&
        (parentFormat & 0xfff) == 0xfff)
    {
      // Something is screwy with the parent format - set to zero
      parentFormat = 0;
      logger.warn("Invalid parent format found - ignoring");
    }


    int alignMask = IntegerHelper.getInt(data[6], data[7]);

    // Get the wrap
    if ((alignMask & 0x08) != 0)
    {
      wrap = true;
    }

    // Get the horizontal alignment
    align = Alignment.getAlignment(alignMask & 0x7);

    // Get the vertical alignment
    valign = VerticalAlignment.getAlignment((alignMask >> 4) & 0x7);

    // Get the orientation
    orientation = Orientation.getOrientation((alignMask >> 8) & 0xff);

    int attr = IntegerHelper.getInt(data[8], data[9]);

    // Get the indentation
    indentation = attr & 0x0F;

    // Get the shrink to fit flag
    shrinkToFit = (attr & 0x10) != 0;

    // Get the used attribute
    if (biffType == biff8)
    {
      usedAttributes = data[9];
    }

    // Get the borders
    int borderMask = IntegerHelper.getInt(data[10], data[11]);

    leftBorder   = BorderLineStyle.getStyle(borderMask & 0x7);
    rightBorder  = BorderLineStyle.getStyle((borderMask >> 4) & 0x7);
    topBorder    = BorderLineStyle.getStyle((borderMask >> 8) & 0x7);
    bottomBorder = BorderLineStyle.getStyle((borderMask >> 12) & 0x7);

    int borderColourMask = IntegerHelper.getInt(data[12], data[13]);

    leftBorderColour = Colour.getInternalColour(borderColourMask & 0x7f);
    rightBorderColour = Colour.getInternalColour
      ((borderColourMask & 0x3f80) >> 7);

    borderColourMask = IntegerHelper.getInt(data[14], data[15]);
    topBorderColour = Colour.getInternalColour(borderColourMask & 0x7f);
    bottomBorderColour = Colour.getInternalColour
      ((borderColourMask & 0x3f80) >> 7);
    
    if (biffType == biff8)
    {
      // Get the background pattern.  This is the six most significant bits
      int patternVal = IntegerHelper.getInt(data[16], data[17]);
      patternVal = patternVal & 0xfc00;
      patternVal = patternVal >> 10;
      pattern = Pattern.getPattern(patternVal);

      // Get the background colour
      int colourPaletteMask = IntegerHelper.getInt(data[18], data[19]);
      backgroundColour = Colour.getInternalColour(colourPaletteMask & 0x3f);

      if (backgroundColour == Colour.UNKNOWN ||
          backgroundColour == Colour.DEFAULT_BACKGROUND1)
      {
        backgroundColour = Colour.DEFAULT_BACKGROUND;
      }
    }
    else
    {
      pattern = Pattern.NONE;
      backgroundColour = Colour.DEFAULT_BACKGROUND;
    }

    // Set the lazy initialization flag
    formatInfoInitialized = true;
  }

  /**
   * Standard hash code implementation
   * @return the hash code
   */
  public int hashCode()
  {
    // Must have its formats info initialized in order to compute the hash code
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    int hashValue = 17;
    int oddPrimeNumber = 37;

    // The boolean fields
    hashValue = oddPrimeNumber*hashValue + (hidden ? 1:0);
    hashValue = oddPrimeNumber*hashValue + (locked ? 1:0);
    hashValue = oddPrimeNumber*hashValue + (wrap ? 1:0);
    hashValue = oddPrimeNumber*hashValue + (shrinkToFit ? 1:0);

    // The enumerations
    if (xfFormatType == cell)
    {
      hashValue = oddPrimeNumber*hashValue + 1;
    }
    else if (xfFormatType == style)
    {
      hashValue = oddPrimeNumber*hashValue + 2;
    }

    hashValue = oddPrimeNumber*hashValue + (align.getValue() + 1);
    hashValue = oddPrimeNumber*hashValue + (valign.getValue() + 1);
    hashValue = oddPrimeNumber*hashValue + (orientation.getValue());

    hashValue ^= leftBorder.getDescription().hashCode();
    hashValue ^= rightBorder.getDescription().hashCode();
    hashValue ^= topBorder.getDescription().hashCode();
    hashValue ^= bottomBorder.getDescription().hashCode();

    hashValue = oddPrimeNumber*hashValue + (leftBorderColour.getValue());
    hashValue = oddPrimeNumber*hashValue + (rightBorderColour.getValue());
    hashValue = oddPrimeNumber*hashValue + (topBorderColour.getValue());
    hashValue = oddPrimeNumber*hashValue + (bottomBorderColour.getValue());
    hashValue = oddPrimeNumber*hashValue + (backgroundColour.getValue());
    hashValue = oddPrimeNumber*hashValue + (pattern.getValue() + 1);

    // The integer fields
    hashValue = oddPrimeNumber*hashValue + usedAttributes;
    hashValue = oddPrimeNumber*hashValue + parentFormat;
    hashValue = oddPrimeNumber*hashValue + fontIndex;
    hashValue = oddPrimeNumber*hashValue + formatIndex;
    hashValue = oddPrimeNumber*hashValue + indentation;

    return hashValue;
  }

  /**
   * Equals method.  This is called when comparing writable formats
   * in order to prevent duplicate formats being added to the workbook
   *
   * @param o object to compare
   * @return TRUE if the objects are equal, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof XFRecord))
    {
      return false;
    }

    XFRecord xfr = (XFRecord) o;

    // Both records must be writable and have their format info initialized
    if (!formatInfoInitialized)
    {
      initializeFormatInformation();
    }

    if (!xfr.formatInfoInitialized)
    {
      xfr.initializeFormatInformation();
    }

    if (xfFormatType   != xfr.xfFormatType ||
        parentFormat   != xfr.parentFormat ||
        locked         != xfr.locked ||
        hidden         != xfr.hidden ||
        usedAttributes != xfr.usedAttributes)
    {
      return false;
    }

    if (align       != xfr.align ||
        valign      != xfr.valign ||
        orientation != xfr.orientation ||
        wrap        != xfr.wrap ||
        shrinkToFit != xfr.shrinkToFit ||
        indentation != xfr.indentation)
    {
      return false;
    }

    if (leftBorder   != xfr.leftBorder  ||
        rightBorder  != xfr.rightBorder ||
        topBorder    != xfr.topBorder   ||
        bottomBorder != xfr.bottomBorder)
    {
      return false;
    }

    if (leftBorderColour   != xfr.leftBorderColour  ||
        rightBorderColour  != xfr.rightBorderColour ||
        topBorderColour    != xfr.topBorderColour   ||
        bottomBorderColour != xfr.bottomBorderColour)
    {
      return false;
    }

    if (backgroundColour != xfr.backgroundColour ||
        pattern          != xfr.pattern)
    {
      return false;
    }

    if (initialized && xfr.initialized)
    {
      // Both formats are initialized, so it is sufficient to just do 
      // shallow equals on font, format objects,
      // since we are testing for the presence of clones anwyay
      // Use indices rather than objects because of the rationalization
      // process (which does not set the object on an XFRecord)
      if (fontIndex   != xfr.fontIndex ||
          formatIndex != xfr.formatIndex)
      {
        return false;
      }
    }
    else
    {
      // Perform a deep compare of fonts and formats
      if (!font.equals(xfr.font) ||
          !format.equals(xfr.format))
      {
        return false;
      }
    }

    return true;
  }

  /**
   * Sets the format index.  This is called during the rationalization process
   * when some of the duplicate number formats have been removed
   * @param newindex the new format index
   */
  void setFormatIndex(int newindex)
  {
    formatIndex = newindex;
  }

  /**
   * Accessor for the font index.  Called by the FormattingRecords objects
   * during the rationalization process
   * @return the font index
   */
  public int getFontIndex()
  {
    return fontIndex;
  }


  /**
   * Sets the font index.  This is called during the rationalization process
   * when some of the duplicate fonts have been removed
   * @param newindex the new index
   */
  void setFontIndex(int newindex)
  {
    fontIndex = newindex;
  }

  /**
   * Sets the format type and parent format from the writable subclass
   * @param t the xf type
   * @param pf the parent format
   */
  protected void setXFDetails(XFType t, int pf)
  {
    xfFormatType = t;
    parentFormat = pf;
  }

  /**
   * Changes the appropriate indexes during the rationalization process
   * @param xfMapping the xf index re-mappings
   */
  void rationalize(IndexMapping xfMapping)
  {
    xfIndex = xfMapping.getNewIndex(xfIndex);

    if (xfFormatType == cell)
    {
      parentFormat = xfMapping.getNewIndex(parentFormat);
    }
  }

  /**
   * Sets the font object with a workbook specific clone.  Called from 
   * the CellValue object when the font has been identified as a statically
   * shared font
   * Also called to superimpose a HyperlinkFont on an existing label cell
   */
  public void setFont(FontRecord f)
  {
    // This style cannot be initialized, otherwise it would mean it would
    // have been initialized with shared font
    // However, sometimes (when setting a row or column format) an initialized
    // XFRecord may have its font overridden by the column/row 

    font = f;
  }
}

