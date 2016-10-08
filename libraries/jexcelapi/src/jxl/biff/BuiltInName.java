/*********************************************************************
*
*      Copyright (C) 2006 Andrew Khan
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
 * Enumeration of built in names
 */
public class BuiltInName
{
  /**
   * The name
   */
  private String name;

  /**
   * The value
   */
  private int value;

  /**
   * The list of name
   */
  private static BuiltInName[] builtInNames = new BuiltInName[0];
  /** 
   * Constructor
   */
  private BuiltInName(String n, int v)
  {
    name = n;
    value = v;

    BuiltInName[] oldnames = builtInNames;
    builtInNames = new BuiltInName[oldnames.length + 1];
    System.arraycopy(oldnames, 0, builtInNames, 0, oldnames.length);
    builtInNames[oldnames.length] = this;
  }

  /**
   * Accessor for the name
   *
   * @return the name
   */
  public String getName()
  {
    return name;
  }

  /**
   * Accessor for the value
   *
   * @return the value
   */
  public int getValue()
  {
    return value;
  }

  /**
   * Gets the built in name for the value
   */
  public static BuiltInName getBuiltInName(int val)
  {
    BuiltInName ret = FILTER_DATABASE;
    for (int i = 0 ; i < builtInNames.length; i++)
    {
      if (builtInNames[i].getValue() == val)
      {
        ret = builtInNames[i];
      }
    }
    return ret;
  }

  // The list of built in names
  public static final BuiltInName CONSOLIDATE_AREA = 
    new BuiltInName("Consolidate_Area", 0x0);
  public static final BuiltInName AUTO_OPEN = 
    new BuiltInName("Auto_Open", 0x1);
  public static final BuiltInName AUTO_CLOSE = 
    new BuiltInName("Auto_Open", 0x2);
  public static final BuiltInName EXTRACT = 
    new BuiltInName("Extract", 0x3);
  public static final BuiltInName DATABASE = 
    new BuiltInName("Database", 0x4);
  public static final BuiltInName CRITERIA = 
    new BuiltInName("Criteria", 0x5);
  public static final BuiltInName PRINT_AREA = 
    new BuiltInName("Print_Area", 0x6);
  public static final BuiltInName PRINT_TITLES = 
    new BuiltInName("Print_Titles", 0x7);
  public static final BuiltInName RECORDER = 
    new BuiltInName("Recorder", 0x8);
  public static final BuiltInName DATA_FORM = 
    new BuiltInName("Data_Form", 0x9);
  public static final BuiltInName AUTO_ACTIVATE = 
    new BuiltInName("Auto_Activate", 0xa);
  public static final BuiltInName AUTO_DEACTIVATE = 
    new BuiltInName("Auto_Deactivate", 0xb);
  public static final BuiltInName SHEET_TITLE = 
    new BuiltInName("Sheet_Title", 0xb);
  public static final BuiltInName FILTER_DATABASE = 
    new BuiltInName("_FilterDatabase", 0xd);
}
