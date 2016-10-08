/*********************************************************************
*
*      Copyright (C) 2005 Andrew Khan
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

import jxl.WorkbookSettings;

/**
 * Helper to get the Microsoft encoded URL from the given string
 */
public class EncodedURLHelper
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(EncodedURLHelper.class);

  // The control codes
  private static byte msDosDriveLetter = 0x01;
  private static byte sameDrive = 0x02;
  private static byte endOfSubdirectory = 0x03;
  private static byte parentDirectory = 0x04;
  private static byte unencodedUrl = 0x05;

  public static byte[] getEncodedURL(String s, WorkbookSettings ws)
  {
    if (s.startsWith("http:"))
    {
      return getURL(s, ws);
    }
    else
    {
      return getFile(s, ws);
    }
  }

  private static byte[] getFile(String s, WorkbookSettings ws)
  {
    ByteArray byteArray = new ByteArray();

    int pos = 0;
    if (s.charAt(1) == ':')
    {
      // we have a drive letter
      byteArray.add(msDosDriveLetter);
      byteArray.add((byte) s.charAt(0));
      pos = 2;
    }
    else if (s.charAt(pos) == '\\' ||
             s.charAt(pos) == '/')
    {
      byteArray.add(sameDrive);
    }

    while (s.charAt(pos) == '\\' ||
           s.charAt(pos) == '/')
    {
      pos++;
    }

    while (pos < s.length())
    {
      int nextSepIndex1 = s.indexOf('/', pos);
      int nextSepIndex2 = s.indexOf('\\', pos);
      int nextSepIndex = 0;
      String nextFileNameComponent = null;

      if (nextSepIndex1 != -1 && nextSepIndex2 != -1)
      {
        // choose the smallest (ie. nearest) separator
        nextSepIndex = Math.min(nextSepIndex1, nextSepIndex2);
      }
      else if (nextSepIndex1 == -1 || nextSepIndex2 == -1)
      {
        // chose the maximum separator
        nextSepIndex = Math.max(nextSepIndex1, nextSepIndex2);
      }

      if (nextSepIndex == -1)
      {
        // no more separators
        nextFileNameComponent = s.substring(pos);
        pos = s.length();
      }
      else
      {
        nextFileNameComponent = s.substring(pos, nextSepIndex);
        pos = nextSepIndex + 1;
      }

      if (nextFileNameComponent.equals("."))
      {
        // current directory - do nothing
      }
      else if (nextFileNameComponent.equals(".."))
      {
        // parent directory
        byteArray.add(parentDirectory);
      }
      else
      {
        // add the filename component
        byteArray.add(StringHelper.getBytes(nextFileNameComponent,
                                            ws));
      }

      if (pos < s.length())
      {
        byteArray.add(endOfSubdirectory);
      }
    }

    return byteArray.getBytes();
  }
  
  private static byte[] getURL(String s, WorkbookSettings ws)
  {
    ByteArray byteArray = new ByteArray();
    byteArray.add(unencodedUrl);
    byteArray.add((byte) s.length());
    byteArray.add(StringHelper.getBytes(s, ws));
    return byteArray.getBytes();
  }
}
