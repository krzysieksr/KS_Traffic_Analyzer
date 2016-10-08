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

/**
 * An enumeration class which contains the  biff types
 */
public final class Type
{
  /**
   * The biff value for this type
   */
  public final int value;
  /**
   * An array of all types
   */
  private static Type[] types = new Type[0];

  /**
   * Constructor
   * Sets the biff value and adds this type to the array of all types
   *
   * @param v the biff code for the type
   */
  private Type(int v)
  {
    value = v;

    // Add to the list of available types
    Type[] newTypes = new Type[types.length + 1];
    System.arraycopy(types, 0, newTypes, 0, types.length);
    newTypes[types.length] = this;
    types = newTypes;
  }

  private static class ArbitraryType {};
  private static ArbitraryType arbitrary = new ArbitraryType();

  /**
   * Constructor used for the creation of arbitrary types
   */
  private Type(int v, ArbitraryType arb)
  {
    value = v;
  }

  /**
   * Standard hash code method
   * @return the hash code
   */
  public int hashCode()
  {
    return value;
  }

  /**
   * Standard equals method
   * @param o the object to compare
   * @return TRUE if the objects are equal, FALSE otherwise
   */
  public boolean equals(Object o)
  {
    if (o == this)
    {
      return true;
    }

    if (!(o instanceof Type))
    {
      return false;
    }

    Type t = (Type) o;

    return value == t.value;
  }

  /**
   * Gets the type object from its integer value
   * @param v the internal code
   * @return the type
   */
  public static Type getType(int v)
  {
    for (int i = 0; i < types.length; i++)
    {
      if (types[i].value == v)
      {
        return types[i];
      }
    }

    return UNKNOWN;
  }

  /**
   * Used to create an arbitrary record type.  This method is only
   * used during bespoke debugging process.  The creation of an
   * arbitrary type does not add it to the static list of known types
   */
  public static Type createType(int v)
  {
    return new Type(v, arbitrary);
  }

  /**
   */
  public static final Type BOF = new Type(0x809);
  /**
   */
  public static final Type EOF = new Type(0x0a);
  /**
   */
  public static final Type BOUNDSHEET = new Type(0x85);
  /**
   */
  public static final Type SUPBOOK = new Type(0x1ae);
  /**
   */
  public static final Type EXTERNSHEET = new Type(0x17);
  /**
   */
  public static final Type DIMENSION = new Type(0x200);
  /**
   */
  public static final Type BLANK = new Type(0x201);
  /**
   */
  public static final Type MULBLANK = new Type(0xbe);
  /**
   */
  public static final Type ROW = new Type(0x208);
  /**
   */
  public static final Type NOTE = new Type(0x1c);
  /**
   */
  public static final Type TXO = new Type(0x1b6);
  /**
   */
  public static final Type RK  = new Type(0x7e);
  /**
   */
  public static final Type RK2  = new Type(0x27e);
  /**
   */
  public static final Type MULRK  = new Type(0xbd);
  /**
   */
  public static final Type INDEX = new Type(0x20b);
  /**
   */
  public static final Type DBCELL = new Type(0xd7);
  /**
   */
  public static final Type SST = new Type(0xfc);
  /**
   */
  public static final Type COLINFO = new Type(0x7d);
  /**
   */
  public static final Type EXTSST = new Type(0xff);
  /**
   */
  public static final Type CONTINUE = new Type(0x3c);
  /**
   */
  public static final Type LABEL = new Type(0x204);
  /**
   */
  public static final Type RSTRING = new Type(0xd6);
  /**
   */
  public static final Type LABELSST = new Type(0xfd);
  /**
   */
  public static final Type NUMBER = new Type(0x203);
  /**
   */
  public static final Type NAME = new Type(0x18);
  /**
   */
  public static final Type TABID = new Type(0x13d);
  /**
   */
  public static final Type ARRAY = new Type(0x221);
  /**
   */
  public static final Type STRING = new Type(0x207);
  /**
   */
  public static final Type FORMULA = new Type(0x406);
  /**
   */
  public static final Type FORMULA2 = new Type(0x6);
  /**
   */
  public static final Type SHAREDFORMULA = new Type(0x4bc);
  /**
   */
  public static final Type FORMAT = new Type(0x41e);
  /**
   */
  public static final Type XF = new Type(0xe0);
  /**
   */
  public static final Type BOOLERR = new Type(0x205);
  /**
   */
  public static final Type INTERFACEHDR = new Type(0xe1);
  /**
   */
  public static final Type SAVERECALC = new Type(0x5f);
  /**
   */
  public static final Type INTERFACEEND = new Type(0xe2);
  /**
   */
  public static final Type XCT = new Type(0x59);
  /**
   */
  public static final Type CRN = new Type(0x5a);
  /**
   */
  public static final Type DEFCOLWIDTH = new Type(0x55);
  /**
   */
  public static final Type DEFAULTROWHEIGHT = new Type(0x225);
  /**
   */
  public static final Type WRITEACCESS = new Type(0x5c);
  /**
   */
  public static final Type WSBOOL = new Type(0x81);
  /**
   */
  public static final Type CODEPAGE = new Type(0x42);
  /**
   */
  public static final Type DSF = new Type(0x161);
  /**
   */
  public static final Type FNGROUPCOUNT = new Type(0x9c);
  /**
   */
  public static final Type FILTERMODE = new Type(0x9b);
  /**
   */
  public static final Type AUTOFILTERINFO = new Type(0x9d);
  /**
   */
  public static final Type AUTOFILTER = new Type(0x9e);
  /**
   */
  public static final Type COUNTRY = new Type(0x8c);
  /**
   */
  public static final Type PROTECT = new Type(0x12);
  /**
   */
  public static final Type SCENPROTECT = new Type(0xdd);
  /**
   */
  public static final Type OBJPROTECT = new Type(0x63);
  /**
   */
  public static final Type PRINTHEADERS = new Type(0x2a);
  /**
   */
  public static final Type HEADER = new Type(0x14);
  /**
   */
  public static final Type FOOTER = new Type(0x15);
  /**
   */
  public static final Type HCENTER = new Type(0x83);
  /**
   */
  public static final Type VCENTER = new Type(0x84);
  /**
   */
  public static final Type FILEPASS = new Type(0x2f);
  /**
   */
  public static final Type SETUP = new Type(0xa1);
  /**
   */
  public static final Type PRINTGRIDLINES = new Type(0x2b);
  /**
   */
  public static final Type GRIDSET = new Type(0x82);
  /**
   */
  public static final Type GUTS = new Type(0x80);
  /**
   */
  public static final Type WINDOWPROTECT = new Type(0x19);
  /**
   */
  public static final Type PROT4REV = new Type(0x1af);
  /**
   */
  public static final Type PROT4REVPASS = new Type(0x1bc);
  /**
   */
  public static final Type PASSWORD = new Type(0x13);
  /**
   */
  public static final Type REFRESHALL = new Type(0x1b7);
  /**
   */
  public static final Type WINDOW1 = new Type(0x3d);
  /**
   */
  public static final Type WINDOW2 = new Type(0x23e);
  /**
   */
  public static final Type BACKUP = new Type(0x40);
  /**
   */
  public static final Type HIDEOBJ = new Type(0x8d);
  /**
   */
  public static final Type NINETEENFOUR = new Type(0x22);
  /**
   */
  public static final Type PRECISION = new Type(0xe);
  /**
   */
  public static final Type BOOKBOOL = new Type(0xda);
  /**
   */
  public static final Type FONT = new Type(0x31);
  /**
   */
  public static final Type MMS = new Type(0xc1);
  /**
   */
  public static final Type CALCMODE = new Type(0x0d);
  /**
   */
  public static final Type CALCCOUNT = new Type(0x0c);
  /**
   */
  public static final Type REFMODE = new Type(0x0f);
  /**
   */
  public static final Type TEMPLATE = new Type(0x60);
  /**
   */
  public static final Type OBJPROJ = new Type(0xd3);
  /**
   */
  public static final Type DELTA = new Type(0x10);
  /**
   */
  public static final Type MERGEDCELLS = new Type(0xe5);
  /**
   */
  public static final Type ITERATION = new Type(0x11);
  /**
   */
  public static final Type STYLE = new Type(0x293);
  /**
   */
  public static final Type USESELFS = new Type(0x160);
  /**
   */
  public static final Type VERTICALPAGEBREAKS = new Type(0x1a);
  /**
   */
  public static final Type HORIZONTALPAGEBREAKS = new Type(0x1b);
  /**
   */
  public static final Type SELECTION = new Type(0x1d);
  /**
   */
  public static final Type HLINK = new Type(0x1b8);
  /**
   */
  public static final Type OBJ = new Type(0x5d);
  /**
   */
  public static final Type MSODRAWING = new Type(0xec);
  /**
   */
  public static final Type MSODRAWINGGROUP = new Type(0xeb);
  /**
   */
  public static final Type LEFTMARGIN = new Type(0x26);
  /**
   */
  public static final Type RIGHTMARGIN = new Type(0x27);
  /**
   */
  public static final Type TOPMARGIN = new Type(0x28);
  /**
   */
  public static final Type BOTTOMMARGIN = new Type(0x29);
  /**
   */
  public static final Type EXTERNNAME = new Type(0x23);
  /**
   */
  public static final Type PALETTE = new Type(0x92);
  /**
   */
  public static final Type PLS = new Type(0x4d);
  /**
   */
  public static final Type SCL = new Type(0xa0);
  /**
   */
  public static final Type PANE = new Type(0x41);
  /**
   */
  public static final Type WEIRD1 = new Type(0xef);
  /**
   */
  public static final Type SORT = new Type(0x90);
  /**
   */
  public static final Type CONDFMT = new Type(0x1b0);
  /**
   */
  public static final Type CF = new Type(0x1b1);
  /**
   */
  public static final Type DV = new Type(0x1be);
  /**
   */
  public static final Type DVAL = new Type(0x1b2);
  /**
   */
  public static final Type BUTTONPROPERTYSET = new Type(0x1ba);
  /**
   *
   */
  public static final Type EXCEL9FILE = new Type(0x1c0);

  // Chart types
  /**
   */
  public static final Type FONTX = new Type(0x1026);
  /**
   */
  public static final Type IFMT = new Type(0x104e);
  /**
   */
  public static final Type FBI = new Type(0x1060);
  /**
   */
  public static final Type ALRUNS = new Type(0x1050);
  /**
   */
  public static final Type SERIES = new Type(0x1003);
  /**
   */
  public static final Type SERIESLIST = new Type(0x1016);
  /**
   */
  public static final Type SBASEREF = new Type(0x1048);
  /**
   */
  public static final Type UNKNOWN = new Type(0xffff);

  // Pivot stuff
  /**
   */
  // public static final Type R = new Type(0xffff);

  // Unknown types
  public static final Type U1C0 = new Type(0x1c0);
  public static final Type U1C1 = new Type(0x1c1);

}









