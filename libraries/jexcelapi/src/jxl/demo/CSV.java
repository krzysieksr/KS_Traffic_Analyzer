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

package jxl.demo;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

/**
 * Simple demo class which uses the api to present the contents
 * of an excel 97 spreadsheet as comma separated values, using a workbook
 * and output stream of your choice
 */
public class CSV
{
  /**
   * Constructor
   *
   * @param w The workbook to interrogate
   * @param out The output stream to which the CSV values are written
   * @param encoding The encoding used by the output stream.  Null or 
   * unrecognized values cause the encoding to default to UTF8
   * @param hide Suppresses hidden cells
   * @exception java.io.IOException
   */
  public CSV(Workbook w, OutputStream out, String encoding, boolean hide)
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
      
      for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++)
      {
        Sheet s = w.getSheet(sheet);

        if (!(hide && s.getSettings().isHidden()))
        {
          bw.write("*** " + s.getName() + " ****");
          bw.newLine();
          
          Cell[] row = null;
          
          for (int i = 0 ; i < s.getRows() ; i++)
          {
            row = s.getRow(i);
            
            if (row.length > 0)
            {
              if (!(hide && row[0].isHidden()))
              {
                bw.write(row[0].getContents());
                // Java 1.4 code to handle embedded commas
                // bw.write("\"" + row[0].getContents().replaceAll("\"","\"\"") + "\"");
              }
              
              for (int j = 1; j < row.length; j++)
              {
                bw.write(',');
                if (!(hide && row[j].isHidden()))
                {
                  bw.write(row[j].getContents());
                  // Java 1.4 code to handle embedded quotes
                  //  bw.write("\"" + row[j].getContents().replaceAll("\"","\"\"") + "\"");
                }
              }
            }
            bw.newLine();
          }
        }
      }
      bw.flush();
      bw.close();
    }
    catch (UnsupportedEncodingException e)
    {
      System.err.println(e.toString());
    }
  }
}



