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

package jxl.biff.drawing;

import java.util.Arrays;

/**
 * Enumeration for the various chunk types
 */
class ChunkType
{
  private byte[] id;
  private String name;

  private static ChunkType[] chunkTypes = new ChunkType[0];

  private ChunkType(int d1, int d2, int d3, int d4, String n)
  {
    id = new byte[] {(byte) d1, (byte) d2, (byte) d3, (byte) d4};
    name = n;

    ChunkType[] ct = new ChunkType[chunkTypes.length + 1];
    System.arraycopy(chunkTypes, 0, ct, 0, chunkTypes.length);
    ct[chunkTypes.length] = this;
    chunkTypes = ct;
  }
 
  public String getName()
  {
    return name;
  }

  public static ChunkType getChunkType(byte d1, byte d2, byte d3, byte d4)
  {
    byte[] cmp = new byte[] {d1, d2, d3, d4};

    boolean found = false;
    ChunkType chunk = ChunkType.UNKNOWN;

    for (int i = 0; i < chunkTypes.length && !found ; i++)
    {
      if (Arrays.equals(chunkTypes[i].id, cmp))
      {
        chunk = chunkTypes[i];
        found = true;
      }
    }

    return chunk;
  }
  

  public static ChunkType IHDR = new ChunkType(0x49, 0x48, 0x44, 0x52,"IHDR");
  public static ChunkType IEND = new ChunkType(0x49, 0x45, 0x4e, 0x44,"IEND");
  public static ChunkType PHYS = new ChunkType(0x70, 0x48, 0x59, 0x73,"pHYs");
  public static ChunkType UNKNOWN = new ChunkType(0xff, 0xff, 0xff, 0xff, "UNKNOWN");
}
