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
import java.util.ArrayList;
import java.util.Iterator;

import jxl.common.Assert;

import jxl.write.biff.File;

/**
 * A container for the list of fonts used in this workbook
 */
public class Fonts
{
  /**
   * The list of fonts
   */
  private ArrayList fonts;

  /**
   * The default number of fonts
   */
  private static final int numDefaultFonts = 4;

  /**
   * Constructor
   */
  public Fonts()
  {
    fonts = new ArrayList();
  }

  /**
   * Adds a font record to this workbook.  If the FontRecord passed in has not
   * been initialized, then its font index is determined based upon the size
   * of the fonts list.  The FontRecord's initialized method is called, and
   * it is added to the list of fonts.
   *
   * @param f the font to add
   */
  public void addFont(FontRecord f)
  {
    if (!f.isInitialized())
    {
      int pos = fonts.size();

      // Remember that the pos with index 4 is skipped
      if (pos >= 4)
      {
        pos++;
      }

      f.initialize(pos);
      fonts.add(f);
    }
  }

  /**
   * Used by FormattingRecord for retrieving the fonts for the
   * hardcoded styles
   *
   * @param index the index of the font to return
   * @return the font with the specified font index
   */
  public FontRecord getFont(int index)
  {
    // remember to allow for the fact that font index 4 is not used
    if (index > 4)
    {
      index--;
    }

    return (FontRecord) fonts.get(index);
  }

  /**
   * Writes out the list of fonts
   *
   * @param outputFile the compound file to write the data to
   * @exception IOException
   */
  public void write(File outputFile) throws IOException
  {
    Iterator i = fonts.iterator();

    FontRecord font = null;
    while (i.hasNext())
    {
      font = (FontRecord) i.next();
      outputFile.write(font);
    }
  }

  /**
   * Rationalizes all the fonts, removing any duplicates
   *
   * @return the mappings between new indexes and old ones
   */
  IndexMapping rationalize()
  {
    IndexMapping mapping = new IndexMapping(fonts.size() + 1);
      // allow for skipping record 4

    ArrayList newfonts = new ArrayList();
    FontRecord fr = null;
    int numremoved = 0;

    // Preserve the default fonts
    for (int i = 0; i < numDefaultFonts; i++)
    {
      fr = (FontRecord) fonts.get(i);
      newfonts.add(fr);
      mapping.setMapping(fr.getFontIndex(), fr.getFontIndex());
    }

    // Now do the rest
    Iterator it = null;
    FontRecord fr2 = null;
    boolean duplicate = false;
    for (int i = numDefaultFonts; i < fonts.size(); i++)
    {
      fr = (FontRecord) fonts.get(i);

      // Compare to all the fonts currently on the list
      duplicate = false;
      it = newfonts.iterator();
      while (it.hasNext() && !duplicate)
      {
        fr2 = (FontRecord) it.next();
        if (fr.equals(fr2))
        {
          duplicate = true;
          mapping.setMapping(fr.getFontIndex(),
                             mapping.getNewIndex(fr2.getFontIndex()));
          numremoved++;
        }
      }

      if (!duplicate)
      {
        // Add to the new list
        newfonts.add(fr);
        int newindex = fr.getFontIndex() - numremoved;
        Assert.verify(newindex > 4);
        mapping.setMapping(fr.getFontIndex(), newindex);
      }
    }

    // Iterate through the remaining fonts, updating all the font indices
    it = newfonts.iterator();
    while (it.hasNext())
    {
      fr = (FontRecord) it.next();
      fr.initialize(mapping.getNewIndex(fr.getFontIndex()));
    }

    fonts = newfonts;

    return mapping;
  }
}
