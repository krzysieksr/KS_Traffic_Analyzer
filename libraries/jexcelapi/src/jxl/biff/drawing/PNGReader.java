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

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;

public class PNGReader
{
  private byte[] pngData;

  private Chunk ihdr;
  private Chunk phys;

  private int pixelWidth;
  private int pixelHeight;
  private int verticalResolution;
  private int horizontalResolution;
  private int resolutionUnit;

  private static byte[] PNG_MAGIC_NUMBER = new byte[]
    {(byte) 0x89, (byte) 0x50, (byte) 0x4e, (byte) 0x47,
     (byte) 0x0d, (byte) 0x0a, (byte) 0x1a, (byte) 0x0a};

  public PNGReader(byte[] data)
  {
    pngData = data;
  }

  void read()
  {
    // Verify the magic data
    byte[] header = new byte[PNG_MAGIC_NUMBER.length];
    System.arraycopy(pngData, 0, header, 0, header.length);
    boolean pngFile = Arrays.equals(PNG_MAGIC_NUMBER, header);
    if (!pngFile)
    {
      return;
    }
    
    int pos = 8;
    while (pos < pngData.length)
    {
      int length = getInt(pngData[pos],
                          pngData[pos+1],
                          pngData[pos+2],
                          pngData[pos+3]);
      ChunkType chunkType = ChunkType.getChunkType(pngData[pos+4],
                                                   pngData[pos+5],
                                                   pngData[pos+6],
                                                   pngData[pos+7]);

      if (chunkType == ChunkType.IHDR)
      {
        ihdr = new Chunk(pos + 8, length, chunkType, pngData);
      }
      else if (chunkType == ChunkType.PHYS)
      {
        phys = new Chunk(pos + 8, length, chunkType, pngData);
      }

      pos += length + 12;
    }

    // Get the width and height from the ihdr
    byte[] ihdrData = ihdr.getData();
    pixelWidth = getInt(ihdrData[0], ihdrData[1], ihdrData[2], ihdrData[3]);
    pixelHeight = getInt(ihdrData[4], ihdrData[5], ihdrData[6], ihdrData[7]);

    if (phys != null)
    {
      byte[] physData = phys.getData();
      resolutionUnit = physData[8];
      horizontalResolution = getInt(physData[0], physData[1], 
                                      physData[2], physData[3]);
      verticalResolution = getInt(physData[4], physData[5], 
                                    physData[6], physData[7]);
    }
  }

  // Gets the big-Endian integer
  private int getInt(byte d1, byte d2, byte d3, byte d4)
  {
    int i1 = d1 & 0xff;
    int i2 = d2 & 0xff;
    int i3 = d3 & 0xff;
    int i4 = d4 & 0xff;

    int val = i1 << 24 |
              i2 << 16 |
              i3 << 8  |
              i4;

    return val;
  }

  public int getHeight()
  {
    return pixelHeight;
  }

  public int getWidth()
  {
    return pixelWidth;
  }

  public int getHorizontalResolution()
  {
    // only return if the resolution unit is in metres
    return resolutionUnit == 1 ? horizontalResolution : 0;
  }

  public int getVerticalResolution()
  {
    // only return if the resolution unit is in metres
    return resolutionUnit == 1 ? verticalResolution : 0;
  }

  public static void main(String args[])
  {
    try
    {
      File f = new File(args[0]);
      int size = (int) f.length();

      byte[] data = new byte[size];

      FileInputStream fis = new FileInputStream(f);
      fis.read(data);
      fis.close();
      PNGReader reader = new PNGReader(data);
      reader.read();
    }
    catch (Throwable t)
    {
      t.printStackTrace();
    }
  }
}
