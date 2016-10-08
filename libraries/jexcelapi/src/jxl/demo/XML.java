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
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.Pattern;

/**
 * Simple demo class which uses the api to present the contents
 * of an excel 97 spreadsheet as an XML document, using a workbook
 * and output stream of your choice
 */
public class XML
{
  /**
   * The output stream to write to
   */
  private OutputStream out;

  /** 
   * The encoding to write
   */
  private String encoding;

  /**
   * The workbook we are reading from
   */
  private Workbook workbook;
  
  /**
   * Constructor
   *
   * @param w The workbook to interrogate
   * @param out The output stream to which the XML values are written
   * @param enc The encoding used by the output stream.  Null or 
   * unrecognized values cause the encoding to default to UTF8
   * @param f Indicates whether the generated XML document should contain
   * the cell format information
   * @exception java.io.IOException
   */
  public XML(Workbook w, OutputStream out, String enc, boolean f)
    throws IOException
  {
    encoding = enc;
    workbook = w;
    this.out = out;

    if (encoding == null || !encoding.equals("UnicodeBig"))
    {
      encoding = "UTF8";
    }

    if (f)
    {
      writeFormattedXML();
    }
    else
    {
      writeXML();
    }

  }

  /**
   * Writes out the workbook data as XML, without formatting information
   */
  private void writeXML() throws IOException
  {
    try
    {
      OutputStreamWriter osw = new OutputStreamWriter(out, encoding);
      BufferedWriter bw = new BufferedWriter(osw);
      
      bw.write("<?xml version=\"1.0\" ?>");
      bw.newLine();
      bw.write("<!DOCTYPE workbook SYSTEM \"workbook.dtd\">");
      bw.newLine();
      bw.newLine();
      bw.write("<workbook>");
      bw.newLine();
      for (int sheet = 0; sheet < workbook.getNumberOfSheets(); sheet++)
      {
        Sheet s = workbook.getSheet(sheet);

        bw.write("  <sheet>");
        bw.newLine();
        bw.write("    <name><![CDATA["+s.getName()+"]]></name>");
        bw.newLine();
      
        Cell[] row = null;
      
        for (int i = 0 ; i < s.getRows() ; i++)
        {
          bw.write("    <row number=\"" + i + "\">");
          bw.newLine();
          row = s.getRow(i);

          for (int j = 0 ; j < row.length; j++)
          {
            if (row[j].getType() != CellType.EMPTY)
            {
              bw.write("      <col number=\"" + j + "\">");
              bw.write("<![CDATA["+row[j].getContents()+"]]>");
              bw.write("</col>");
              bw.newLine();
            }
          }
          bw.write("    </row>");
          bw.newLine();
        }
        bw.write("  </sheet>");
        bw.newLine();
      }
      
      bw.write("</workbook>");
      bw.newLine();

      bw.flush();
      bw.close();
    }
    catch (UnsupportedEncodingException e)
    {
      System.err.println(e.toString());
    }
  }

  /**
   * Writes out the workbook data as XML, with formatting information
   */
  private void writeFormattedXML() throws IOException
  {
    try
    {
      OutputStreamWriter osw = new OutputStreamWriter(out, encoding);
      BufferedWriter bw = new BufferedWriter(osw);
      
      bw.write("<?xml version=\"1.0\" ?>");
      bw.newLine();
      bw.write("<!DOCTYPE workbook SYSTEM \"formatworkbook.dtd\">");
      bw.newLine();
      bw.newLine();
      bw.write("<workbook>");
      bw.newLine();
      for (int sheet = 0; sheet < workbook.getNumberOfSheets(); sheet++)
      {
        Sheet s = workbook.getSheet(sheet);

        bw.write("  <sheet>");
        bw.newLine();
        bw.write("    <name><![CDATA["+s.getName()+"]]></name>");
        bw.newLine();
      
        Cell[] row = null;
        CellFormat format = null;
        Font font = null;
      
        for (int i = 0 ; i < s.getRows() ; i++)
        {
          bw.write("    <row number=\"" + i + "\">");
          bw.newLine();
          row = s.getRow(i);

          for (int j = 0 ; j < row.length; j++)
          {
            // Remember that empty cells can contain format information
            if ((row[j].getType() != CellType.EMPTY) ||
                (row[j].getCellFormat() != null))
            {
              format = row[j].getCellFormat();
              bw.write("      <col number=\"" + j + "\">");
              bw.newLine();
              bw.write("        <data>");
              bw.write("<![CDATA["+row[j].getContents()+"]]>");
              bw.write("</data>");
              bw.newLine();     

              if (row[j].getCellFormat() != null)
              {
                bw.write("        <format wrap=\"" + format.getWrap() + "\"");
                bw.newLine();
                bw.write("                align=\"" + 
                         format.getAlignment().getDescription() + "\"");
                bw.newLine();
                bw.write("                valign=\"" + 
                         format.getVerticalAlignment().getDescription() + "\"");
                bw.newLine();
                bw.write("                orientation=\"" + 
                         format.getOrientation().getDescription() + "\"");
                bw.write(">");
                bw.newLine();

                // The font information
                font = format.getFont();
                bw.write("          <font name=\"" + font.getName() + "\"");
                bw.newLine();
                bw.write("                point_size=\"" + 
                         font.getPointSize() + "\"");
                bw.newLine();
                bw.write("                bold_weight=\"" + 
                         font.getBoldWeight() + "\"");
                bw.newLine();
                bw.write("                italic=\"" + font.isItalic() + "\"");
                bw.newLine();
                bw.write("                underline=\"" + 
                         font.getUnderlineStyle().getDescription() + "\"");
                bw.newLine();
                bw.write("                colour=\"" + 
                         font.getColour().getDescription() + "\"");
                bw.newLine();
                bw.write("                script=\"" + 
                         font.getScriptStyle().getDescription() + "\"");
                bw.write(" />");
                bw.newLine();


                // The cell background information
                if (format.getBackgroundColour() != Colour.DEFAULT_BACKGROUND ||
                    format.getPattern()          != Pattern.NONE)
                {
                  bw.write("          <background colour=\"" + 
                           format.getBackgroundColour().getDescription() + "\"");
                  bw.newLine();
                  bw.write("                      pattern=\"" + 
                           format.getPattern().getDescription() + "\"");
                  bw.write(" />");
                  bw.newLine();
                }


                // The cell border, if it has one
                if (format.getBorder(Border.TOP  )  != BorderLineStyle.NONE ||
                    format.getBorder(Border.BOTTOM) != BorderLineStyle.NONE ||
                    format.getBorder(Border.LEFT)   != BorderLineStyle.NONE ||
                    format.getBorder(Border.RIGHT)  != BorderLineStyle.NONE)
                {
                  
                  bw.write("          <border top=\"" + 
                           format.getBorder(Border.TOP).getDescription() + "\"");
                  bw.newLine();
                  bw.write("                  bottom=\"" + 
                           format.getBorder(Border.BOTTOM).getDescription() + 
                           "\"");
                  bw.newLine();
                  bw.write("                  left=\"" + 
                           format.getBorder(Border.LEFT).getDescription() + "\"");
                  bw.newLine();
                  bw.write("                  right=\"" + 
                           format.getBorder(Border.RIGHT).getDescription() + "\"");
                  bw.write(" />");
                  bw.newLine();
                }

                // The cell number/date format
                if (!format.getFormat().getFormatString().equals(""))
                {
                  bw.write("          <format_string string=\"");
                  bw.write(format.getFormat().getFormatString());
                  bw.write("\" />");
                  bw.newLine();
                }

                bw.write("        </format>");
                bw.newLine();
              }

              bw.write("      </col>");
              bw.newLine();
            }
          }
          bw.write("    </row>");
          bw.newLine();
        }
        bw.write("  </sheet>");
        bw.newLine();
      }
      
      bw.write("</workbook>");
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






