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

package jxl.demo;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;

import jxl.Workbook;
import jxl.biff.drawing.DrawingGroup;
import jxl.biff.drawing.EscherDisplay;
import jxl.read.biff.WorkbookParser;

/**
 * Displays the escher data
 */
public class EscherDrawingGroup
{
  /**
   * Constructor
   *
   * @param w The workbook to interrogate
   * @param out The output stream to which the CSV values are written
   * @param encoding The encoding used by the output stream.  Null or 
   * unrecognized values cause the encoding to default to UTF8
   * @exception java.io.IOException
   */
  public EscherDrawingGroup (Workbook w, OutputStream out, String encoding)
    throws IOException
  {
    if (encoding == null || !encoding.equals("UnicodeBig"))
    {
      encoding = "UTF8";
    }

    try
    {
      OutputStreamWriter osw = new OutputStreamWriter(out, encoding);
      BufferedWriter bw = new BufferedWriter(osw);

      WorkbookParser wp = (WorkbookParser) w;

      DrawingGroup dg = wp.getDrawingGroup();
      
      if (dg != null)
      {
        EscherDisplay ed = new EscherDisplay(dg, bw);
        ed.display();
      }
      
      bw.newLine();
      bw.newLine();
      bw.flush();
      bw.close();
    }
    catch (UnsupportedEncodingException e)
    {
      System.err.println(e.toString());
    }
  }

}

