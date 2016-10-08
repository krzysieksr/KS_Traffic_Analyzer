/*********************************************************************
 *
 *      Copyright (C) 2004 Andrew Khan, Eric Jung
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

import jxl.common.Logger;

/**
 * Class which represents an Excel header or footer. Information for this
 * class came from Microsoft Knowledge Base Article 142136 
 * (previously Q142136).
 *
 * This class encapsulates three internal structures representing the header
 * or footer contents which appear on the left, right or central part of the 
 * page
 */
public abstract class HeaderFooter
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(HeaderFooter.class);

  // Codes to format text

  /**
   * Turns bold printing on or off
   */
  private static final String BOLD_TOGGLE = "&B";

  /**
   * Turns underline printing on or off
   */
  private static final String UNDERLINE_TOGGLE = "&U";

  /**
   * Turns italic printing on or off
   */
  private static final String ITALICS_TOGGLE = "&I";

  /**
   * Turns strikethrough printing on or off
   */
  private static final String STRIKETHROUGH_TOGGLE = "&S";

  /**
   * Turns double-underline printing on or off
   */
  private static final String DOUBLE_UNDERLINE_TOGGLE = "&E";

  /**
   * Turns superscript printing on or off
   */
  private static final String SUPERSCRIPT_TOGGLE = "&X";

  /**
   * Turns subscript printing on or off
   */
  private static final String SUBSCRIPT_TOGGLE = "&Y";

  /**
   * Turns outline printing on or off (Macintosh only)
   */
  private static final String OUTLINE_TOGGLE = "&O";

  /**
   * Turns shadow printing on or off (Macintosh only)
   */
  private static final String SHADOW_TOGGLE = "&H";

  /**
   * Left-aligns the characters that follow
   */
  private static final String LEFT_ALIGN = "&L";

  /**
   * Centres the characters that follow
   */
  private static final String CENTRE = "&C";

  /**
   * Right-aligns the characters that follow
   */
  private static final String RIGHT_ALIGN = "&R";
  
  // Codes to insert specific data

  /**
   * Prints the page number
   */
  private static final String PAGENUM = "&P";

  /**
   * Prints the total number of pages in the document
   */
  private static final String TOTAL_PAGENUM = "&N";

  /**
   * Prints the current date
   */
  private static final String DATE = "&D";

  /**
   * Prints the current time
   */
  private static final String TIME = "&T";

  /**
   * Prints the name of the workbook
   */
  private static final String WORKBOOK_NAME = "&F";

  /**
   * Prints the name of the worksheet
   */
  private static final String WORKSHEET_NAME = "&A";

  /**
   * The contents - a simple wrapper around a string buffer
   */
  protected static class Contents
  {
    /**
     * The buffer containing the header/footer string
     */
    private StringBuffer contents;

    /**
     * The constructor
     */
    protected Contents()
    {
      contents = new StringBuffer();
    }

    /**
     * Constructor used when reading worksheets.  The string contains all
     * the formatting (but not alignment characters
     *
     * @param s the format string
     */
    protected Contents(String s)
    {
      contents = new StringBuffer(s);
    }

    /**
     * Copy constructor
     *
     * @param copy the contents to copy
     */
    protected Contents(Contents copy)
    {
      contents = new StringBuffer(copy.getContents());
    }

    /**
     * Retrieves a <code>String</code>ified
     * version of this object
     *
     * @return the header string
     */
    protected String getContents()
    {
      return contents != null ? contents.toString() : "";
    }

    /**
     * Internal method which appends the text to the string buffer
     *
     * @param txt
     */
    private void appendInternal(String txt)
    {
      if (contents == null)
      {
        contents = new StringBuffer();
      }

      contents.append(txt);
    }

    /**
     * Internal method which appends the text to the string buffer
     *
     * @param ch
     */
    private void appendInternal(char ch)
    {
      if (contents == null)
      {
        contents = new StringBuffer();
      }

      contents.append(ch);
    }

    /**
     * Appends the text to the string buffer
     *
     * @param txt
     */
    protected void append(String txt)
    {
      appendInternal(txt);
    }
  
    /**
     * Turns bold printing on or off. Bold printing
     * is initially off. Text subsequently appended to
     * this object will be bolded until this method is
     * called again.
     */
    protected void toggleBold()
    {
     appendInternal(BOLD_TOGGLE);
    }

    /**
     * Turns underline printing on or off. Underline printing
     * is initially off. Text subsequently appended to
     * this object will be underlined until this method is
     * called again.
     */
    protected void toggleUnderline()
    {
      appendInternal(UNDERLINE_TOGGLE);
    }

    /**
     * Turns italics printing on or off. Italics printing
     * is initially off. Text subsequently appended to
     * this object will be italicized until this method is
     * called again.
     */
    protected void toggleItalics()
    {
      appendInternal(ITALICS_TOGGLE);
    }

    /**
     * Turns strikethrough printing on or off. Strikethrough printing
     * is initially off. Text subsequently appended to
     * this object will be striked out until this method is
     * called again.
     */
    protected void toggleStrikethrough()
    {
      appendInternal(STRIKETHROUGH_TOGGLE);
    }  

    /**
     * Turns double-underline printing on or off. Double-underline printing
     * is initially off. Text subsequently appended to
     * this object will be double-underlined until this method is
     * called again.
     */
    protected void toggleDoubleUnderline()
    {
      appendInternal(DOUBLE_UNDERLINE_TOGGLE);
    }      

    /**
     * Turns superscript printing on or off. Superscript printing
     * is initially off. Text subsequently appended to
     * this object will be superscripted until this method is
     * called again.
     */
    protected void toggleSuperScript()
    {
      appendInternal(SUPERSCRIPT_TOGGLE);
    } 	

    /**
     * Turns subscript printing on or off. Subscript printing
     * is initially off. Text subsequently appended to
     * this object will be subscripted until this method is
     * called again.
     */
    protected void toggleSubScript()
    {
      appendInternal(SUBSCRIPT_TOGGLE);
    }   
    
    /**
     * Turns outline printing on or off (Macintosh only).
     * Outline printing is initially off. Text subsequently appended
     * to this object will be outlined until this method is
     * called again.
     */
    protected void toggleOutline()
    {
     appendInternal(OUTLINE_TOGGLE);
    }
    
    /**
     * Turns shadow printing on or off (Macintosh only).
     * Shadow printing is initially off. Text subsequently appended
     * to this object will be shadowed until this method is
     * called again.
     */
    protected void toggleShadow()
    {
      appendInternal(SHADOW_TOGGLE);
    }
    
    /**
     * Sets the font of text subsequently appended to this
     * object.. Previously appended text is not affected.
     * <p/>
     * <strong>Note:</strong> no checking is performed to
     * determine if fontName is a valid font.
     * 
     * @param fontName name of the font to use
     */
    protected void setFontName(String fontName)
    {
      // Font name must be in quotations
      appendInternal("&\"");
      appendInternal(fontName);
      appendInternal('\"');
    }

    /**
     * Sets the font size of text subsequently appended to this
     * object. Previously appended text is not affected.
     * <p/>
     * Valid point sizes are between 1 and 99 (inclusive). If
     * size is outside this range, this method returns false
     * and does not change font size. If size is within this
     * range, the font size is changed and true is returned. 
     * 
     * @param size The size in points. Valid point sizes are
     * between 1 and 99 (inclusive).
     * @return true if the font size was changed, false if font
     * size was not changed because 1 > size > 99. 
     */
    protected boolean setFontSize(int size)
    {
      if (size < 1 || size > 99)
      {
        return false;
      }
  	
      // A two digit number should be used -- even if the
      // leading number is just a zero.
      String fontSize;
      if (size < 10) 
      {
  	  // single-digit -- make two digit
        fontSize = "0" + size; 
      }
      else
      {
        fontSize = Integer.toString(size);
      }
  	
      appendInternal('&');
      appendInternal(fontSize);
      return true;
    }
  
    /**
     * Appends the page number
     */
    protected void appendPageNumber()
    {
      appendInternal(PAGENUM);
    }
  
    /**
     * Appends the total number of pages
     */
    protected void appendTotalPages()
    {
      appendInternal(TOTAL_PAGENUM);
    }
  
    /**
     * Appends the current date
     */
    protected void appendDate()
    {
      appendInternal(DATE); 
    }
  
    /**
     * Appends the current time
     */
    protected void appendTime()
    {
      appendInternal(TIME);
    }
  
    /**
     * Appends the workbook name
     */
    protected void appendWorkbookName()
    {
      appendInternal(WORKBOOK_NAME);
    }
    
    /**
     * Appends the worksheet name
     */
    protected void appendWorkSheetName()
    {
      appendInternal(WORKSHEET_NAME);
    }

    /**
     * Clears the contents of this portion
     */
    protected void clear()
    {
      contents = null;
    }

    /**
     * Queries if the contents are empty
     *
     * @return TRUE if the contents are empty, FALSE otherwise
     */
    protected boolean empty()
    {
      if (contents == null || contents.length() == 0)
      {
        return true;
      }
      else
      {
        return false;
      }
    }
  }

  /**
   * The left aligned header/footer contents
   */
  private Contents left;

  /**
   * The right aligned header/footer contents
   */
  private Contents right;

  /**
   * The centrally aligned header/footer contents
   */
  private Contents centre;

  /**
   * Default constructor.
   */
  protected HeaderFooter()
  {
    left = createContents();
    right = createContents();
    centre = createContents();
  }

  /**
   * Copy constructor
   *
   * @param c the item to copy
   */
  protected HeaderFooter(HeaderFooter hf)
  {
    left = createContents(hf.left);
    right = createContents(hf.right);
    centre = createContents(hf.centre);
  }

  /**
   * Constructor used when reading workbooks to separate the left, right
   * a central part of the strings into their constituent parts
   */
  protected HeaderFooter(String s)
  {
    if (s == null || s.length() == 0)
    {
      left = createContents();
      right = createContents();
      centre = createContents();
      return;
    }

    int leftPos = s.indexOf(LEFT_ALIGN);
    int rightPos = s.indexOf(RIGHT_ALIGN);
    int centrePos = s.indexOf(CENTRE);

    if (leftPos == -1 && rightPos == -1 && centrePos == -1)
    {
      // When no part is specified, it is the center part
      centre = createContents(s);
    }
    else
    {
      // Left part?
      if (leftPos != -1)
      {
        // We have a left part, find end of left part
        int endLeftPos= s.length();
        if (centrePos > leftPos)
        {
          // Case centre part behind left part
          endLeftPos= centrePos;
          if (rightPos > leftPos && endLeftPos > rightPos)
          {
            // LRC case
            endLeftPos = rightPos;
          }
          else
          {
            // LCR case
          }
        }
        else
        {
          // Case centre part before left part
          if (rightPos > leftPos)
          {
            // LR case
            endLeftPos= rightPos;
          }
          else
          {
            // *L case
            // Left pos is last


          }
        }
        left = createContents(s.substring(leftPos + 2, endLeftPos));
      }

      // Right part?
      if (rightPos != -1)
      {
        // Find end of right part
        int endRightPos= s.length();
        if (centrePos > rightPos)
        {
          // centre part behind right part
          endRightPos= centrePos;
          if (leftPos > rightPos && endRightPos > leftPos)
          {
            // RLC case
            endRightPos= leftPos;
          }
          else
          {
            // RCL case
          }
        }
        else
        {
          if (leftPos > rightPos)
          {
            // RL case
            endRightPos= leftPos;
          }
          else
          {
            // *R case
            // Right pos is last
          }
        }
        right = createContents(s.substring(rightPos + 2, endRightPos));
      }

      // Centre part?
      if (centrePos != -1)
      {
        // Find end of centre part
        int endCentrePos= s.length();
        if (rightPos > centrePos)
        {
          // right part behind centre part
          endCentrePos= rightPos;
          if (leftPos > centrePos && endCentrePos > leftPos)
          {
            // CLR case
            endCentrePos= leftPos;
          }
          else
          {
            // CRL case
          }
        }
        else
        {
          if (leftPos > centrePos)
          {
            // CL case
            endCentrePos= leftPos;
          }
          else
          {
            // *C case
            // Centre pos is last
          }
        }
        centre = createContents(s.substring(centrePos + 2, endCentrePos));
      }
    }


    if (left == null)
    {
      left = createContents();
    }

    if (centre == null)
    {
      centre = createContents();
    }

    if (right == null)
    {
      right = createContents();
    }
  }

  /**
   * Retrieves a <code>String</code>ified
   * version of this object
   *
   * @return the header string
   */
  public String toString()
  {
    StringBuffer hf = new StringBuffer();
    if (!left.empty())
    {
      hf.append(LEFT_ALIGN);
      hf.append(left.getContents());
    }

    if (!centre.empty())
    {
      hf.append(CENTRE);
      hf.append(centre.getContents());
    }

    if (!right.empty())
    {
      hf.append(RIGHT_ALIGN);
      hf.append(right.getContents());
    }

    return hf.toString();
  }

  /**
   * Accessor for the contents which appear on the right hand side of the page
   *
   * @return the right aligned contents
   */
  protected Contents getRightText()
  {
    return right;
  }

  /**
   * Accessor for the contents which in the centre of the page
   *
   * @return the centrally  aligned contents
   */
  protected Contents getCentreText()
  {
    return centre;
  }

  /**
   * Accessor for the contents which appear on the left hand side of the page
   *
   * @return the left aligned contents
   */
  protected Contents getLeftText()
  {
    return left;
  }

  /**
   * Clears the contents of the header/footer
   */
  protected void clear()
  {
    left.clear();
    right.clear();
    centre.clear();
  }

  /**
   * Creates internal class of the appropriate type
   */
  protected abstract Contents createContents();

  /**
   * Creates internal class of the appropriate type
   */
  protected abstract Contents createContents(String s);

  /**
   * Creates internal class of the appropriate type
   */
  protected abstract Contents createContents(Contents c);
}
