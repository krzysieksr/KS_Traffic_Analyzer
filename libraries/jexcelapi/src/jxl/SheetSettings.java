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

package jxl;

import jxl.common.Assert;

import jxl.format.PageOrientation;
import jxl.format.PaperSize;
import jxl.biff.SheetRangeImpl;
import jxl.Range;
import jxl.format.PageOrder;

/**
 * This is a bean which client applications may use to get/set various
 * properties which are associated with a particular worksheet, such
 * as headers and footers, page orientation etc.
 */
public final class SheetSettings
{
  /**
   * The page orientation
   */
  private PageOrientation orientation;
  
  /**
   * The page order
   */
  private PageOrder pageOrder;

  /**
   * The paper size for printing
   */
  private PaperSize paperSize;

  /**
   * Indicates whether or not this sheet is protected
   */
  private boolean sheetProtected;

  /**
   * Indicates whether or not this sheet is hidden
   */
  private boolean hidden;

  /**
   * Indicates whether or not this sheet is selected
   */
  private boolean selected;

  /**
   * The header
   */
  private HeaderFooter header;

  /**
   * The margin allocated for any page headers, in inches
   */
  private double headerMargin;

  /**
   * The footer
   */
  private HeaderFooter footer;

  /**
   * The margin allocated for any page footers, in inches
   */
  private double footerMargin;

  /**
   * The scale factor used when printing
   */
  private int scaleFactor;

  /**
   * The zoom factor used when viewing.  Note the difference between
   * this and the scaleFactor which is used when printing
   */
  private int zoomFactor;

  /**
   * The page number at which to commence printing
   */
  private int pageStart;

  /**
   * The number of pages into which this excel sheet is squeezed widthwise
   */
  private int fitWidth;

  /**
   * The number of pages into which this excel sheet is squeezed heightwise
   */
  private int fitHeight;

  /**
   * The horizontal print resolution
   */
  private int horizontalPrintResolution;

  /**
   * The vertical print resolution
   */
  private int verticalPrintResolution;

  /**
   * The margin from the left hand side of the paper in inches
   */
  private double leftMargin;

  /**
   * The margin from the right hand side of the paper in inches
   */
  private double rightMargin;

  /**
   * The margin from the top of the paper in inches
   */
  private double topMargin;

  /**
   * The margin from the bottom of the paper in inches
   */
  private double bottomMargin;

  /**
   * Indicates whether to fit the print to the pages or scale the output
   * This field is manipulated indirectly by virtue of the setFitWidth/Height
   * methods
   */
  private boolean fitToPages;

  /**
   * Indicates whether grid lines should be displayed
   */
  private boolean showGridLines;

  /**
   * Indicates whether grid lines should be printed
   */
  private boolean printGridLines;

  /**
   * Indicates whether sheet headings should be printed
   */
  private boolean printHeaders;

  /**
   * Indicates the view mode
   */
  private boolean pageBreakPreviewMode;

  /**
   * Indicates whether the sheet should display zero values
   */
  private boolean displayZeroValues;

  /**
   * The password for protected sheets
   */
  private String password;

  /**
   * The password hashcode - used when copying sheets
   */
  private int passwordHash;

  /**
   * The default column width, in characters
   */
  private int defaultColumnWidth;

  /**
   * The default row height, in 1/20th of a point
   */
  private int defaultRowHeight;

  /**
   * The horizontal freeze pane
   */
  private int horizontalFreeze;

  /**
   * The vertical freeze position
   */
  private int verticalFreeze;

  /**
   * Vertical centre flag
   */
  private boolean verticalCentre;

  /**
   * Horizontal centre flag
   */
  private boolean horizontalCentre;

  /**
   * The number of copies to print
   */
  private int copies;

  /**
   * Automatic formula calculation
   */
  private boolean automaticFormulaCalculation;

  /**
   * Recalculate the formulas before save
   */
  private boolean recalculateFormulasBeforeSave;

  /**
   * The magnification factor for use during page break preview mode (in 
   * percent)
   */
  private int pageBreakPreviewMagnification;

  /**
   * The magnification factor for use during normal mode (in percent)
   */
  private int normalMagnification;

  /**
   * The print area
   */
  private Range printArea;

  /**
   * The print row titles
   */
  private Range printTitlesRow;

  /**
   * The print column titles
   */
  private Range printTitlesCol;
    
  /**
   * A handle to the sheet - used internally for ranges
   */
  private Sheet sheet;

  // ***
  // The defaults
  // **
  private static final PageOrientation DEFAULT_ORIENTATION =
    PageOrientation.PORTRAIT;
  private static final PageOrder DEFAULT_ORDER =
    PageOrder.RIGHT_THEN_DOWN;
  private static final PaperSize DEFAULT_PAPER_SIZE = PaperSize.A4;
  private static final double DEFAULT_HEADER_MARGIN = 0.5;
  private static final double DEFAULT_FOOTER_MARGIN = 0.5;
  private static final int    DEFAULT_PRINT_RESOLUTION = 0x12c;
  private static final double DEFAULT_WIDTH_MARGIN     = 0.75;
  private static final double DEFAULT_HEIGHT_MARGIN    = 1;

  private static final int DEFAULT_DEFAULT_COLUMN_WIDTH = 8;
  private static final int DEFAULT_ZOOM_FACTOR = 100;
  private static final int DEFAULT_NORMAL_MAGNIFICATION = 100;
  private static final int DEFAULT_PAGE_BREAK_PREVIEW_MAGNIFICATION = 60;

  // The publicly accessible values
  /**
   * The default value for the default row height
   */
  public static final int DEFAULT_DEFAULT_ROW_HEIGHT = 0xff;

  /**
   * Default constructor
   */
  public SheetSettings(Sheet s)
  {
    sheet = s; // for internal use, when accessing ranges
    orientation        = DEFAULT_ORIENTATION;
    pageOrder          = DEFAULT_ORDER;
    paperSize          = DEFAULT_PAPER_SIZE;
    sheetProtected     = false;
    hidden             = false;
    selected           = false;
    headerMargin       = DEFAULT_HEADER_MARGIN;
    footerMargin       = DEFAULT_FOOTER_MARGIN;
    horizontalPrintResolution = DEFAULT_PRINT_RESOLUTION;
    verticalPrintResolution   = DEFAULT_PRINT_RESOLUTION;
    leftMargin         = DEFAULT_WIDTH_MARGIN;
    rightMargin        = DEFAULT_WIDTH_MARGIN;
    topMargin          = DEFAULT_HEIGHT_MARGIN;
    bottomMargin       = DEFAULT_HEIGHT_MARGIN;
    fitToPages         = false;
    showGridLines      = true;
    printGridLines     = false;
    printHeaders       = false;
    pageBreakPreviewMode = false;
    displayZeroValues  = true;
    defaultColumnWidth = DEFAULT_DEFAULT_COLUMN_WIDTH;
    defaultRowHeight   = DEFAULT_DEFAULT_ROW_HEIGHT;
    zoomFactor         = DEFAULT_ZOOM_FACTOR;
    pageBreakPreviewMagnification = DEFAULT_PAGE_BREAK_PREVIEW_MAGNIFICATION;
    normalMagnification = DEFAULT_NORMAL_MAGNIFICATION;
    horizontalFreeze   = 0;
    verticalFreeze     = 0;
    copies             = 1;
    header             = new HeaderFooter();
    footer             = new HeaderFooter();
    automaticFormulaCalculation = true;
    recalculateFormulasBeforeSave = true;
  }

  /**
   * Copy constructor.  Called when copying sheets
   * @param copy the settings to copy
   */
  public SheetSettings(SheetSettings copy, Sheet s)
  {
    Assert.verify(copy != null);

    sheet = s;  // for internal use when accessing ranges
    orientation    = copy.orientation;
    pageOrder      = copy.pageOrder;
    paperSize      = copy.paperSize;
    sheetProtected = copy.sheetProtected;
    hidden         = copy.hidden;
    selected       = false; // don't copy the selected flag
    headerMargin   = copy.headerMargin;
    footerMargin   = copy.footerMargin;
    scaleFactor    = copy.scaleFactor;
    pageStart      = copy.pageStart;
    fitWidth       = copy.fitWidth;
    fitHeight      = copy.fitHeight;
    horizontalPrintResolution = copy.horizontalPrintResolution;
    verticalPrintResolution   = copy.verticalPrintResolution;
    leftMargin         = copy.leftMargin;
    rightMargin        = copy.rightMargin;
    topMargin          = copy.topMargin;
    bottomMargin       = copy.bottomMargin;
    fitToPages         = copy.fitToPages;
    password           = copy.password;
    passwordHash       = copy.passwordHash;
    defaultColumnWidth = copy.defaultColumnWidth;
    defaultRowHeight   = copy.defaultRowHeight;
    zoomFactor         = copy.zoomFactor;
    pageBreakPreviewMagnification = copy.pageBreakPreviewMagnification;
    normalMagnification = copy.normalMagnification;
    showGridLines      = copy.showGridLines;
    displayZeroValues  = copy.displayZeroValues;
    pageBreakPreviewMode = copy.pageBreakPreviewMode;
    horizontalFreeze   = copy.horizontalFreeze;
    verticalFreeze     = copy.verticalFreeze;
    horizontalCentre   = copy.horizontalCentre;
    verticalCentre     = copy.verticalCentre;
    copies             = copy.copies;
    header             = new HeaderFooter(copy.header);
    footer             = new HeaderFooter(copy.footer);
    automaticFormulaCalculation = copy.automaticFormulaCalculation;
    recalculateFormulasBeforeSave = copy.recalculateFormulasBeforeSave;

    if (copy.printArea != null)
    {
      printArea = new SheetRangeImpl
        (sheet,
         copy.getPrintArea().getTopLeft().getColumn(),
         copy.getPrintArea().getTopLeft().getRow(),
         copy.getPrintArea().getBottomRight().getColumn(),
         copy.getPrintArea().getBottomRight().getRow());
    }
    
    if (copy.printTitlesRow != null)
    {
      printTitlesRow = new SheetRangeImpl
        (sheet,
         copy.getPrintTitlesRow().getTopLeft().getColumn(),
         copy.getPrintTitlesRow().getTopLeft().getRow(),
         copy.getPrintTitlesRow().getBottomRight().getColumn(),
         copy.getPrintTitlesRow().getBottomRight().getRow());
    }    

    if (copy.printTitlesCol != null)
    {
      printTitlesCol = new SheetRangeImpl
        (sheet,
         copy.getPrintTitlesCol().getTopLeft().getColumn(),
         copy.getPrintTitlesCol().getTopLeft().getRow(),
         copy.getPrintTitlesCol().getBottomRight().getColumn(),
         copy.getPrintTitlesCol().getBottomRight().getRow());
    }    
  }

  /**
   * Sets the paper orientation for printing this sheet
   *
   * @param po the orientation
   */
  public void setOrientation(PageOrientation po)
  {
    orientation = po;
  }

  /**
   * Accessor for the orientation
   *
   * @return the orientation
   */
  public PageOrientation getOrientation()
  {
    return orientation;
  }

  /**
   * Accessor for the order
   * 
   * @return
   */
  public PageOrder getPageOrder() 
  {
    return pageOrder;
  }

  /**
   * Sets the page order for printing this sheet
   * 
   * @param order
   */
  public void setPageOrder(PageOrder order) 
  {
    this.pageOrder = order;
  }

  /**
   * Sets the paper size to be used when printing this sheet
   *
   * @param ps the paper size
   */
  public void setPaperSize(PaperSize ps)
  {
    paperSize = ps;
  }

  /**
   * Accessor for the paper size
   *
   * @return the paper size
   */
  public PaperSize getPaperSize()
  {
    return paperSize;
  }

  /**
   * Queries whether this sheet is protected (ie. read only)
   *
   * @return TRUE if this sheet is read only, FALSE otherwise
   */
  public boolean isProtected()
  {
    return sheetProtected;
  }

  /**
   * Sets the protected (ie. read only) status of this sheet
   *
   * @param p the protected status
   */
  public void setProtected(boolean p)
  {
    sheetProtected = p;
  }

  /**
   * Sets the margin for any page headers
   *
   * @param d the margin in inches
   */
  public void setHeaderMargin(double d)
  {
    headerMargin = d;
  }

  /**
   * Accessor for the header margin
   *
   * @return the header margin
   */
  public double getHeaderMargin()
  {
    return headerMargin;
  }

  /**
   * Sets the margin for any page footer
   *
   * @param d the footer margin in inches
   */
  public void setFooterMargin(double d)
  {
    footerMargin = d;
  }

  /**
   * Accessor for the footer margin
   *
   * @return the footer margin
   */
  public double getFooterMargin()
  {
    return footerMargin;
  }

  /**
   * Sets the hidden status of this worksheet
   *
   * @param h the hidden flag
   */
  public void setHidden(boolean h)
  {
    hidden = h;
  }

  /**
   * Accessor for the hidden nature of this sheet
   *
   * @return TRUE if this sheet is hidden, FALSE otherwise
   */
  public boolean isHidden()
  {
    return hidden;
  }

  /**
   * Sets this sheet to be when it is opened in excel
   *
   * @deprecated use overloaded version which takes a boolean
   */
  public void setSelected()
  {
    setSelected(true);
  }

  /**
   * Sets this sheet to be when it is opened in excel
   *
   * @param s sets whether this sheet is selected or not
   */
  public void setSelected(boolean s)
  {
    selected = s;
  }

  /**
   * Accessor for the selected nature of the sheet
   *
   * @return TRUE if this sheet is selected, FALSE otherwise
   */
  public boolean isSelected()
  {
    return selected;
  }

  /**
   * Sets the scale factor for this sheet to be used when printing.  The
   * parameter is a percentage, therefore setting a scale factor of 100 will
   * print at normal size, 50 half size, 200 double size etc
   *
   * @param sf the scale factor as a percentage
   */
  public void setScaleFactor(int sf)
  {
    scaleFactor = sf;
    fitToPages = false;
  }

  /**
   * Accessor for the scale factor
   *
   * @return the scale factor
   */
  public int getScaleFactor()
  {
    return scaleFactor;
  }

  /**
   * Sets the page number at which to commence printing
   *
   * @param ps the page start number
   */
  public void setPageStart(int ps)
  {
    pageStart = ps;
  }

  /**
   * Accessor for the page start
   *
   * @return the page start
   */
  public int getPageStart()
  {
    return pageStart;
  }

  /**
   * Sets the number of pages widthwise which this sheet should be
   * printed into
   *
   * @param fw the number of pages
   */
  public void setFitWidth(int fw)
  {
    fitWidth = fw;
    fitToPages = true;
  }

  /**
   * Accessor for the fit width
   *
   * @return the number of pages this sheet will be printed into widthwise
   */
  public int getFitWidth()
  {
    return fitWidth;
  }

  /**
   * Sets the number of pages vertically that this sheet will be printed into
   *
   * @param fh the number of pages this sheet will be printed into heightwise
   */
  public void setFitHeight(int fh)
  {
    fitHeight = fh;
    fitToPages = true;
  }

  /**
   * Accessor for the fit height
   *
   * @return the number of pages this sheet will be printed into heightwise
   */
  public int getFitHeight()
  {
    return fitHeight;
  }

  /**
   * Sets the horizontal print resolution
   *
   * @param hpw the print resolution
   */
  public void setHorizontalPrintResolution(int hpw)
  {
    horizontalPrintResolution = hpw;
  }

  /**
   * Accessor for the horizontal print resolution
   *
   * @return the horizontal print resolution
   */
  public int getHorizontalPrintResolution()
  {
    return horizontalPrintResolution;
  }

  /**
   * Sets the vertical print reslution
   *
   * @param vpw the vertical print resolution
   */
  public void setVerticalPrintResolution(int vpw)
  {
    verticalPrintResolution = vpw;
  }

  /**
   * Accessor for the vertical print resolution
   *
   * @return the vertical print resolution
   */
  public int getVerticalPrintResolution()
  {
    return verticalPrintResolution;
  }

  /**
   * Sets the right margin
   *
   * @param m the right margin in inches
   */
  public void setRightMargin(double m)
  {
    rightMargin = m;
  }

  /**
   * Accessor for the right margin
   *
   * @return the right margin in inches
   */
  public double getRightMargin()
  {
    return rightMargin;
  }

  /**
   * Sets the left margin
   *
   * @param m the left margin in inches
   */
  public void setLeftMargin(double m)
  {
    leftMargin = m;
  }

  /**
   * Accessor for the left margin
   *
   * @return the left margin in inches
   */
  public double getLeftMargin()
  {
    return leftMargin;
  }

  /**
   * Sets the top margin
   *
   * @param m the top margin in inches
   */
  public void setTopMargin(double m)
  {
    topMargin = m;
  }

  /**
   * Accessor for the top margin
   *
   * @return the top margin in inches
   */
  public double getTopMargin()
  {
    return topMargin;
  }

  /**
   * Sets the bottom margin
   *
   * @param m the bottom margin in inches
   */
  public void setBottomMargin(double m)
  {
    bottomMargin = m;
  }

  /**
   * Accessor for the bottom margin
   *
   * @return the bottom margin in inches
   */
  public double getBottomMargin()
  {
    return bottomMargin;
  }

  /**
   * Gets the default margin width
   *
   * @return the default margin width
   */
  public double getDefaultWidthMargin()
  {
    return DEFAULT_WIDTH_MARGIN;
  }

  /**
   * Gets the default margin height
   *
   * @return the default margin height
   */
  public double getDefaultHeightMargin()
  {
    return DEFAULT_HEIGHT_MARGIN;
  }

  /**
   * Accessor for the fit width print flag
   * @return TRUE if the print is to fit to pages, false otherwise
   */
  public boolean getFitToPages()
  {
    return fitToPages;
  }

  /**
   * Accessor for the fit to pages flag
   * @param b TRUE to fit to pages, FALSE to use a scale factor
   */
  public void setFitToPages(boolean b)
  {
    fitToPages = b;
  }

  /**
   * Accessor for the password
   *
   * @return the password to unlock this sheet, or NULL if not protected
   */
  public String getPassword()
  {
    return password;
  }

  /**
   * Sets the password for this sheet
   *
   * @param s the password
   */
  public void setPassword(String s)
  {
    password = s;
  }

  /**
   * Accessor for the password hash - used only when copying sheets
   *
   * @return passwordHash
   */
  public int getPasswordHash()
  {
    return passwordHash;
  }

  /**
   * Accessor for the password hash - used only when copying sheets
   *
   * @param ph the password hash
   */
  public void setPasswordHash(int ph)
  {
    passwordHash = ph;
  }

  /**
   * Accessor for the default column width
   *
   * @return the default column width, in characters
   */
  public int getDefaultColumnWidth()
  {
    return defaultColumnWidth;
  }

  /**
   * Sets the default column width
   *
   * @param w the new default column width
   */
  public void setDefaultColumnWidth(int w)
  {
    defaultColumnWidth = w;
  }

  /**
   * Accessor for the default row height
   *
   * @return the default row height, in 1/20ths of a point
   */
  public int getDefaultRowHeight()
  {
    return defaultRowHeight;
  }

  /**
   * Sets the default row height
   *
   * @param h the default row height, in 1/20ths of a point
   */
  public void setDefaultRowHeight(int h)
  {
    defaultRowHeight = h;
  }

  /**
   * Accessor for the zoom factor.  Do not confuse zoom factor (which relates
   * to the on screen view) with scale factor (which refers to the scale factor
   * when printing)
   *
   * @return the zoom factor as a percentage
   */
  public int getZoomFactor()
  {
    return zoomFactor;
  }

  /**
   * Sets the zoom factor.  Do not confuse zoom factor (which relates
   * to the on screen view) with scale factor (which refers to the scale factor
   * when printing)
   *
   * @param zf the zoom factor as a percentage
   */
  public void setZoomFactor(int zf)
  {
    zoomFactor = zf;
  }

  /**
   * Accessor for the page break preview mangificaton factor.
   * Do not confuse zoom factor or scale factor
   *
   * @return the page break preview magnification  a percentage
   */
  public int getPageBreakPreviewMagnification()
  {
    return pageBreakPreviewMagnification;
  }

  /**
   * Accessor for the page break preview magnificaton factor.
   * Do not confuse zoom factor or scale factor
   *
   * @param f the page break preview magnification as a percentage
   */
  public void setPageBreakPreviewMagnification(int f)
  {
    pageBreakPreviewMagnification =f ;
  }

  /**
   * Accessor for the nomral view  magnificaton factor.
   * Do not confuse zoom factor or scale factor
   *
   * @return the page break preview magnification  a percentage
   */
  public int getNormalMagnification()
  {
    return normalMagnification;
  }

  /**
   * Accessor for the normal magnificaton factor.
   * Do not confuse zoom factor or scale factor
   *
   * @param f the page break preview magnification as a percentage
   */
  public void setNormalMagnification(int f)
  {
    normalMagnification = f ;
  }


  /**
   * Accessor for the displayZeroValues property
   *
   * @return TRUE to display zero values, FALSE not to bother
   */
  public boolean getDisplayZeroValues()
  {
    return displayZeroValues;
  }

  /**
   * Sets the displayZeroValues property
   *
   * @param b TRUE to show zero values, FALSE not to bother
   */
  public void setDisplayZeroValues(boolean b)
  {
    displayZeroValues = b;
  }

  /**
   * Accessor for the showGridLines property
   *
   * @return TRUE if grid lines will be shown, FALSE otherwise
   */
  public boolean getShowGridLines()
  {
    return showGridLines;
  }

  /**
   * Sets the showGridLines property
   *
   * @param b TRUE to show grid lines on this sheet, FALSE otherwise
   */
  public void setShowGridLines(boolean b)
  {
    showGridLines = b;
  }

  /**
   * Accessor for the pageBreakPreview mode
   *
   * @return TRUE if page break preview is enabled, FALSE otherwise
   */
  public boolean getPageBreakPreviewMode()
  {
    return pageBreakPreviewMode;
  }

  /**
   * Sets the pageBreakPreviewMode  property
   *
   * @param b TRUE to launch in page break preview mode, FALSE otherwise
   */
  public void setPageBreakPreviewMode(boolean b)
  {
    pageBreakPreviewMode = b;
  }

  /**
   * Accessor for the printGridLines property
   *
   * @return TRUE if grid lines will be printed, FALSE otherwise
   */
  public boolean getPrintGridLines()
  {
    return printGridLines;
  }

   /**
   * Sets the printGridLines property
   *
   * @param b TRUE to print grid lines on this sheet, FALSE otherwise
   */
  public void setPrintGridLines(boolean b)
  {
    printGridLines = b;
  }

  /**
   * Accessor for the printHeaders property
   *
   * @return TRUE if headers will be printed, FALSE otherwise
   */
  public boolean getPrintHeaders()
  {
    return printHeaders;
  }

   /**
   * Sets the printHeaders property
   *
   * @param b TRUE to print headers on this sheet, FALSE otherwise
   */
  public void setPrintHeaders(boolean b)
  {
    printHeaders = b;
  }

  /**
   * Gets the row at which the pane is frozen horizontally
   *
   * @return the row at which the pane is horizontally frozen, or 0 if there
   * is no freeze
   */
  public int getHorizontalFreeze()
  {
    return horizontalFreeze;
  }

  /**
   * Sets the row at which the pane is frozen horizontally
   *
   * @param row the row number to freeze at
   */
  public void setHorizontalFreeze(int row)
  {
    horizontalFreeze = Math.max(row, 0);
  }

  /**
   * Gets the column at which the pane is frozen vertically
   *
   * @return the column at which the pane is vertically frozen, or 0 if there
   * is no freeze
   */
  public int getVerticalFreeze()
  {
    return verticalFreeze;
  }

  /**
   * Sets the row at which the pane is frozen vertically
   *
   * @param col the column number to freeze at
   */
  public void setVerticalFreeze(int col)
  {
    verticalFreeze = Math.max(col, 0);
  }

  /**
   * Sets the number of copies
   *
   * @param c the number of copies
   */
  public void setCopies(int c)
  {
    copies = c;
  }

  /**
   * Accessor for the number of copies to print
   *
   * @return the number of copies
   */
  public int getCopies()
  {
    return copies;
  }

  /**
   * Accessor for the header
   *
   * @return the header
   */
  public HeaderFooter getHeader()
  {
    return header;
  }

  /**
   * Sets the header
   *
   * @param h the header
   */
  public void setHeader(HeaderFooter h)
  {
    header = h;
  }

  /**
   * Sets the footer
   *
   * @param f the footer
   */
  public void setFooter(HeaderFooter f)
  {
    footer = f;
  }

  /**
   * Accessor for the footer
   *
   * @return the footer
   */
  public HeaderFooter getFooter()
  {
    return footer;
  }

  /**
   * Accessor for the horizontal centre
   *
   * @return Returns the horizontalCentre.
   */
  public boolean isHorizontalCentre()
  {
    return horizontalCentre;
  }

  /**
   * Sets the horizontal centre
   *
   * @param horizCentre The horizontalCentre to set.
   */
  public void setHorizontalCentre(boolean horizCentre)
  {
    this.horizontalCentre = horizCentre;
  }

  /**
   * Accessor for the vertical centre
   *
   * @return Returns the verticalCentre.
   */
  public boolean isVerticalCentre()
  {
    return verticalCentre;
  }

  /**
   * Sets the vertical centre
   *
   * @param vertCentre The verticalCentre to set.
   */
  public void setVerticalCentre(boolean vertCentre)
  {
    this.verticalCentre = vertCentre;
  }

  /**
   * Sets the automatic formula calculation flag
   *
   * @param auto - TRUE to automatically calculate the formulas,
   * FALSE otherwise
   */
  public void setAutomaticFormulaCalculation(boolean auto)
  {
    automaticFormulaCalculation = auto;
  }

  /**
   * Retrieves the automatic formula calculation flag
   *
   * @return TRUE if formulas are calculated automatically, FALSE if they
   * are calculated manually
   */
  public boolean getAutomaticFormulaCalculation()
  {
    return automaticFormulaCalculation;
  }

  /**
   * Sets the recalculate formulas when the sheet is saved flag
   *
   * @param recalc - TRUE to automatically calculate the formulas when the,
   * spreadsheet is saved, FALSE otherwise
   */
  public void setRecalculateFormulasBeforeSave(boolean recalc)
  {
    recalculateFormulasBeforeSave = recalc;
  }

  /**
   * Retrieves the recalculate formulas before save  flag
   *
   * @return TRUE if formulas are calculated before the sheet is saved,
   * FALSE otherwise
   */
  public boolean getRecalculateFormulasBeforeSave()
  {
    return recalculateFormulasBeforeSave;
  }

  /**
   * Sets the print area for this sheet
   *
   * @param firstCol the first column of the print area
   * @param firstRow the first row of the print area
   * @param lastCol the last column of the print area
   * @param lastRow the last row of the print area
   */
  public void setPrintArea(int firstCol, 
                           int firstRow,  
                           int lastCol,
                           int lastRow)
  {
    printArea = new SheetRangeImpl(sheet, firstCol, firstRow, 
                                   lastCol, lastRow);
  }

  /**
   * Accessor for the print area
   *
   * @return the print area, or NULL if one is not defined for this sheet
   */
  public Range getPrintArea()
  {
    return printArea;
  }

  /**
   * Sets both of the print titles for this sheet
   *
   * @param firstRow the first row of the print row titles
   * @param lastRow the last row of the print row titles
   * @param firstCol the first column of the print column titles
   * @param lastCol the last column of the print column titles
   */
  public void setPrintTitles(int firstRow, 
                             int lastRow,  
                             int firstCol,
                             int lastCol)
  {
    setPrintTitlesRow(firstRow, lastRow);
    setPrintTitlesCol(firstCol, lastCol);
  }

  /**
   * Sets the print row titles for this sheet
   *
   * @param firstRow the first row of the print titles
   * @param lastRow the last row of the print titles
   */
  public void setPrintTitlesRow(int firstRow, 
                           		int lastRow)
  {
    printTitlesRow = new SheetRangeImpl(sheet, 0, firstRow, 
    									     255, lastRow);
  }
  
  /**
   * Sets the print column titles for this sheet
   *
   * @param firstRow the first row of the print titles
   * @param lastRow the last row of the print titles
   */
  public void setPrintTitlesCol(int firstCol, 
                           		int lastCol)
  {
    printTitlesCol = new SheetRangeImpl(sheet, firstCol, 0, 
                                      		    lastCol, 65535);
  }
  
  /**
   * Accessor for the print row titles
   *
   * @return the print row titles, or NULL if one is not defined for this sheet
   */
  public Range getPrintTitlesRow()
  {
    return printTitlesRow;
  }
  
  /**
   * Accessor for the print column titles
   *
   * @return the print column titles, or NULL if one is not defined for this 
   * sheet
   */
  public Range getPrintTitlesCol()
  {
    return printTitlesCol;
  }
}
