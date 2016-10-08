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

import jxl.common.Assert;
import jxl.common.Logger;

/**
 * Contains the common data for a compound file
 */
public abstract class BaseCompoundFile
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(BaseCompoundFile.class);

  /**
   * The identifier at the beginning of every OLE file
   */
  protected static final byte[] IDENTIFIER = new byte[]
    {(byte) 0xd0,
     (byte) 0xcf,
     (byte) 0x11,
     (byte) 0xe0,
     (byte) 0xa1,
     (byte) 0xb1,
     (byte) 0x1a,
     (byte) 0xe1};
  /**
   */
  protected static final int NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
  /**
   */
  protected static final int SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
  /**
   */
  protected static final int NUM_SMALL_BLOCK_DEPOT_BLOCKS_POS = 0x40;
  /**
   */
  protected static final int ROOT_START_BLOCK_POS = 0x30;
  /**
   */
  protected static final int BIG_BLOCK_SIZE  = 0x200;
  /**
   */
  protected static final int SMALL_BLOCK_SIZE = 0x40;
  /**
   */
  protected static final int EXTENSION_BLOCK_POS = 0x44;
  /**
   */
  protected static final int NUM_EXTENSION_BLOCK_POS = 0x48;
  /**
   */
  protected static final int PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
  /**
   */
  protected static final int BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
  /**
   */
  protected static final int SMALL_BLOCK_THRESHOLD = 0x1000;

  // property storage offsets
    /**
     */
    private static final int SIZE_OF_NAME_POS = 0x40;
    /**
     */
    private static final int TYPE_POS = 0x42;
    /**
    */
    private static final int COLOUR_POS = 0x43;
    /**
     */
    private static final int PREVIOUS_POS = 0x44;
    /**
     */
    private static final int NEXT_POS = 0x48;
    /**
     */
    private static final int CHILD_POS = 0x4c;
    /**
     */
    private static final int START_BLOCK_POS = 0x74;
    /**
     */
    private static final int SIZE_POS = 0x78;

  /**
   * The standard property sets
   */
  public final static String ROOT_ENTRY_NAME = "Root Entry";
  public final static String WORKBOOK_NAME = "Workbook";
  public final static String SUMMARY_INFORMATION_NAME = 
    "\u0005SummaryInformation";
  public final static String DOCUMENT_SUMMARY_INFORMATION_NAME = 
    "\u0005DocumentSummaryInformation";
  public final static String COMP_OBJ_NAME = 
    "\u0001CompObj";
  public final static String[] STANDARD_PROPERTY_SETS  = 
    new String[] {ROOT_ENTRY_NAME, WORKBOOK_NAME,
                  SUMMARY_INFORMATION_NAME,
                  DOCUMENT_SUMMARY_INFORMATION_NAME};

  /**
   * Property storage types
   */
  public final static int NONE_PS_TYPE = 0;
  public final static int DIRECTORY_PS_TYPE = 1;
  public final static int FILE_PS_TYPE = 2;
  public final static int ROOT_ENTRY_PS_TYPE = 5;


  /**
   * Inner class to represent the property storage sets.  Access is public
   * to allow access from the PropertySetsReader demo utility
   */
  public class PropertyStorage
  {
    /**
     * The name of this property set
     */
    public String name;
    /**
     * The type of the property set
     */
    public int type;
    /**
     * The colour of the property set
     */
    public int colour;
    /**
     * The block number in the stream which this property sets starts at
     */
    public int startBlock;
    /**
     * The size, in bytes, of this property set
     */
    public int size;
    /**
     * The previous property set
     */
    public int previous;
    /**
     * The next property set
     */
    public int next;
    /**
     * The child for this property set
     */
    public int child;

    /**
     * The data that created this set
     */
    public byte[] data;

    /**
     * Constructs a property set
     *
     * @param d the bytes
     */
    public PropertyStorage(byte[] d)
    {
      data = d;
      int nameSize = IntegerHelper.getInt(data[SIZE_OF_NAME_POS],
                                          data[SIZE_OF_NAME_POS + 1]);

      if (nameSize > SIZE_OF_NAME_POS)
      {
        logger.warn("property set name exceeds max length - truncating");
        nameSize = SIZE_OF_NAME_POS;

      }
      type = data[TYPE_POS];
      colour = data[COLOUR_POS];

      startBlock = IntegerHelper.getInt
        (data[START_BLOCK_POS],
         data[START_BLOCK_POS + 1],
         data[START_BLOCK_POS + 2],
         data[START_BLOCK_POS + 3]);
      size = IntegerHelper.getInt
        (data[SIZE_POS],
         data[SIZE_POS + 1],
         data[SIZE_POS + 2],
         data[SIZE_POS + 3]);
      previous = IntegerHelper.getInt
        (data[PREVIOUS_POS],
         data[PREVIOUS_POS+1],
         data[PREVIOUS_POS+2],
         data[PREVIOUS_POS+3]);
      next = IntegerHelper.getInt
        (data[NEXT_POS],
         data[NEXT_POS+1],
         data[NEXT_POS+2],
         data[NEXT_POS+3]);
      child = IntegerHelper.getInt
        (data[CHILD_POS],
         data[CHILD_POS+1],
         data[CHILD_POS+2],
         data[CHILD_POS+3]);

      int chars = 0;
      if (nameSize > 2)
      {
        chars = (nameSize - 1) / 2;
      }

      StringBuffer n = new StringBuffer("");
      for (int i = 0; i < chars ; i++)
      {
        n.append( (char) data[i * 2]);
      }

      name = n.toString();
    }

    /**
     * Constructs an empty property set.  Used when writing the file
     *
     * @param name the property storage name
     */
    public PropertyStorage(String name)
    {
      data = new byte[PROPERTY_STORAGE_BLOCK_SIZE];

      Assert.verify(name.length() < 32);

      IntegerHelper.getTwoBytes((name.length() + 1) * 2,
                                data,
                                SIZE_OF_NAME_POS);
      // add one to the name length to allow for the null character at
      // the end
      for (int i = 0; i < name.length(); i++)
      {
        data[i * 2] = (byte) name.charAt(i);
      }
    }

    /**
     * Sets the type
     *
     * @param t the type
     */
    public void setType(int t)
    {
      type = t;
      data[TYPE_POS] = (byte) t;
    }

    /**
     * Sets the number of the start block
     *
     * @param sb the number of the start block
     */
    public void setStartBlock(int sb)
    {
      startBlock = sb;
      IntegerHelper.getFourBytes(sb, data, START_BLOCK_POS);
    }

    /**
     * Sets the size of the file
     *
     * @param s the size
     */
    public void setSize(int s)
    {
      size = s;
      IntegerHelper.getFourBytes(s, data, SIZE_POS);
    }

    /**
     * Sets the previous block
     *
     * @param prev the previous block
     */
    public void setPrevious(int prev)
    {
      previous = prev;
      IntegerHelper.getFourBytes(prev, data, PREVIOUS_POS);
    }

    /**
     * Sets the next block
     *
     * @param nxt the next block
     */
    public void setNext(int nxt)
    {
      next = nxt;
      IntegerHelper.getFourBytes(next, data, NEXT_POS);
    }

    /**
     * Sets the child
     *
     * @param dir the child
     */
    public void setChild(int dir)
    {
      child = dir;
      IntegerHelper.getFourBytes(child, data, CHILD_POS);
    }

    /**
     * Sets the colour
     *
     * @param col colour
     */
    public void setColour(int col)
    {
      colour = col == 0 ? 0 : 1;
      data[COLOUR_POS] = (byte) colour;
    }

  }

  /**
   * Constructor
   */
  protected BaseCompoundFile()
  {
  }

}




