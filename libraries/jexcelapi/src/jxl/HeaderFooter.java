/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan, Eric Jung
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

/**
 * Class which represents an Excel header or footer.
 */
public final class HeaderFooter extends jxl.biff.HeaderFooter
{
  /**
   * The contents - a simple wrapper around a string buffer
   */
  public static class Contents extends jxl.biff.HeaderFooter.Contents
  {
    /**
     * The constructor
     */
    Contents()
    {
      super();
    }

    /**
     * Constructor used when reading worksheets.  The string contains all
     * the formatting (but not alignment characters
     *
     * @param s the format string
     */
    Contents(String s)
    {
      super(s);
    }

    /**
     * Copy constructor
     *
     * @param copy the contents to copy
     */
    Contents(Contents copy)
    {
      super(copy);
    }

    /**
     * Appends the text to the string buffer
     *
     * @param txt the text to append
     */
    public void append(String txt)
    {
      super.append(txt);
    }

    /**
     * Turns bold printing on or off. Bold printing
     * is initially off. Text subsequently appended to
     * this object will be bolded until this method is
     * called again.
     */
    public void toggleBold()
    {
      super.toggleBold();
    }

    /**
     * Turns underline printing on or off. Underline printing
     * is initially off. Text subsequently appended to
     * this object will be underlined until this method is
     * called again.
     */
    public void toggleUnderline()
    {
      super.toggleUnderline();
    }

    /**
     * Turns italics printing on or off. Italics printing
     * is initially off. Text subsequently appended to
     * this object will be italicized until this method is
     * called again.
     */
    public void toggleItalics()
    {
      super.toggleItalics();
    }

    /**
     * Turns strikethrough printing on or off. Strikethrough printing
     * is initially off. Text subsequently appended to
     * this object will be striked out until this method is
     * called again.
     */
    public void toggleStrikethrough()
    {
      super.toggleStrikethrough();
    }

    /**
     * Turns double-underline printing on or off. Double-underline printing
     * is initially off. Text subsequently appended to
     * this object will be double-underlined until this method is
     * called again.
     */
    public void toggleDoubleUnderline()
    {
      super.toggleDoubleUnderline();
    }

    /**
     * Turns superscript printing on or off. Superscript printing
     * is initially off. Text subsequently appended to
     * this object will be superscripted until this method is
     * called again.
     */
    public void toggleSuperScript()
    {
      super.toggleSuperScript();
    }

    /**
     * Turns subscript printing on or off. Subscript printing
     * is initially off. Text subsequently appended to
     * this object will be subscripted until this method is
     * called again.
     */
    public void toggleSubScript()
    {
      super.toggleSubScript();
    }

    /**
     * Turns outline printing on or off (Macintosh only).
     * Outline printing is initially off. Text subsequently appended
     * to this object will be outlined until this method is
     * called again.
     */
    public void toggleOutline()
    {
      super.toggleOutline();
    }

    /**
     * Turns shadow printing on or off (Macintosh only).
     * Shadow printing is initially off. Text subsequently appended
     * to this object will be shadowed until this method is
     * called again.
     */
    public void toggleShadow()
    {
      super.toggleShadow();
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
    public void setFontName(String fontName)
    {
      super.setFontName(fontName);
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
    public boolean setFontSize(int size)
    {
      return super.setFontSize(size);
    }

    /**
     * Appends the page number
     */
    public void appendPageNumber()
    {
      super.appendPageNumber();
    }

    /**
     * Appends the total number of pages
     */
    public void appendTotalPages()
    {
      super.appendTotalPages();
    }

    /**
     * Appends the current date
     */
    public void appendDate()
    {
      super.appendDate();
    }

    /**
     * Appends the current time
     */
    public void appendTime()
    {
      super.appendTime();
    }

    /**
     * Appends the workbook name
     */
    public void appendWorkbookName()
    {
      super.appendWorkbookName();
    }

    /**
     * Appends the worksheet name
     */
    public void appendWorkSheetName()
    {
      super.appendWorkSheetName();
    }

    /**
     * Clears the contents of this portion
     */
    public void clear()
    {
      super.clear();
    }

    /**
     * Queries if the contents are empty
     *
     * @return TRUE if the contents are empty, FALSE otherwise
     */
    public boolean empty()
    {
      return super.empty();
    }
  }

  /**
   * Default constructor.
   */
  public HeaderFooter()
  {
    super();
  }

  /**
   * Copy constructor
   *
   * @param hf the item to copy
   */
  public HeaderFooter(HeaderFooter hf)
  {
    super(hf);
  }

  /**
   * Constructor used when reading workbooks to separate the left, right
   * a central part of the strings into their constituent parts
   *
   * @param s the header string
   */
  public HeaderFooter(String s)
  {
    super(s);
  }

  /**
   * Retrieves a <code>String</code>ified
   * version of this object
   *
   * @return the header string
   */
  public String toString()
  {
    return super.toString();
  }

  /**
   * Accessor for the contents which appear on the right hand side of the page
   *
   * @return the right aligned contents
   */
  public Contents getRight()
  {
    return (Contents) super.getRightText();
  }

  /**
   * Accessor for the contents which in the centre of the page
   *
   * @return the centrally  aligned contents
   */
  public Contents getCentre()
  {
    return (Contents) super.getCentreText();
  }

  /**
   * Accessor for the contents which appear on the left hand side of the page
   *
   * @return the left aligned contents
   */
  public Contents getLeft()
  {
    return (Contents) super.getLeftText();
  }

  /**
   * Clears the contents of the header/footer
   */
  public void clear()
  {
    super.clear();
  }

  /**
   * Creates internal class of the appropriate type
   *
   * @return the created contents
   */
  protected jxl.biff.HeaderFooter.Contents createContents()
  {
    return new Contents();
  }

  /**
   * Creates internal class of the appropriate type
   *
   * @param s the string to create the contents
   * @return the created contents
   */
  protected jxl.biff.HeaderFooter.Contents createContents(String s)
  {
    return new Contents(s);
  }

  /**
   * Creates internal class of the appropriate type
   *
   * @param c the contents to copy
   * @return the new contents
   */
  protected jxl.biff.HeaderFooter.Contents
    createContents(jxl.biff.HeaderFooter.Contents c)
  {
    return new Contents((Contents) c);
  }
}
