/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.IntegerHelper;
import jxl.write.biff.File;

/**
 * Handles the writing out of the different charts and images on a sheet.
 * Called by the SheetWriter object
 */
public class SheetDrawingWriter
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(SheetDrawingWriter.class);

  /**
   * The drawings on the sheet
   */
  private ArrayList drawings;

  /**
   * Flag indicating whether the drawings on the sheet were modified
   */
  private boolean drawingsModified;

  /**
   * The charts on the sheet
   */
  private Chart[] charts;

  /**
   * The workbook settings
   */
  private WorkbookSettings workbookSettings;

  /**
   * Constructor
   *
   * @param ws the workbook settings
   */
  public SheetDrawingWriter(WorkbookSettings ws)
  {
    charts = new Chart[0];
  }

  /**
   * The drawings on the sheet
   *
   * @param dr the list of drawings
   * @param mod flag indicating whether the drawings have been tampered with
   */
  public void setDrawings(ArrayList dr, boolean mod)
  {
    drawings = dr;
    drawingsModified = mod;
  }

  /**
   * Writes out the MsoDrawing records and Obj records for each image
   * and chart on the sheet
   *
   * @param outputFile the output file
   * @exception IOException
   */
  public void write(File outputFile) throws IOException
  {
    // If there are no drawings or charts on this sheet then exit
    if (drawings.size() == 0 && charts.length == 0)
    {
      return;
    }

    // See if any drawing has been modified
    boolean modified = drawingsModified;
    int numImages = drawings.size();

    for (Iterator i = drawings.iterator(); i.hasNext() && !modified;)
    {
      DrawingGroupObject d = (DrawingGroupObject) i.next();
      if (d.getOrigin() != Origin.READ)
      {
        modified = true;
      }
    }

    // If the drawing order has been muddled at all, then we'll need
    // to regenerate the Escher drawing data
    if (numImages > 0 && !modified)
    {
      DrawingGroupObject d2 = (DrawingGroupObject) drawings.get(0);
      if (!d2.isFirst())
      {
        modified = true;
      }
    }

    // Check to see if this sheet consists of just a single chart.  If so
    // there is no MsoDrawingRecord, so write out the data and exit
    if (numImages == 0 &&
        charts.length == 1 &&
        charts[0].getMsoDrawingRecord() == null)
    {
      modified = false; // this sheet has not been modified
    }

    // If no drawing has been modified, then simply write them straight out
    // again and exit
    if (!modified)
    {
      writeUnmodified(outputFile);
      return;
    }

    Object[] spContainerData = new Object[numImages + charts.length];
    int length = 0;
    EscherContainer firstSpContainer = null;

    // Get all the spContainer byte data from the drawings
    // and store in an array
    for (int i = 0; i < numImages; i++)
    {
      DrawingGroupObject drawing = (DrawingGroupObject) drawings.get(i);

      EscherContainer spc = drawing.getSpContainer();

      if (spc != null)
      {
        byte[] data = spc.getData();
        spContainerData[i] = data;

        if (i == 0)
        {
          firstSpContainer = spc;
        }
        else
        {
          length += data.length;
        }
      }
    }

    // Get all the spContainer bytes from the charts and add to the array
    for (int i = 0; i < charts.length; i++)
    {
      EscherContainer spContainer = charts[i].getSpContainer();
      byte[] data = spContainer.getBytes(); //use getBytes instead of getData
      data = spContainer.setHeaderData(data);
      spContainerData[i + numImages] = data;

      if (i == 0 && numImages == 0)
      {
        firstSpContainer = spContainer;
      }
      else
      {
        length += data.length;
      }
    }

    // Put the generalised stuff around the first item
    DgContainer dgContainer = new DgContainer();
    Dg dg  = new Dg(numImages + charts.length);
    dgContainer.add(dg);

    SpgrContainer spgrContainer = new SpgrContainer();

    SpContainer spContainer = new SpContainer();
    Spgr spgr = new Spgr();
    spContainer.add(spgr);
    Sp sp = new Sp(ShapeType.MIN, 1024, 5);
    spContainer.add(sp);
    spgrContainer.add(spContainer);

    spgrContainer.add(firstSpContainer);

    dgContainer.add(spgrContainer);

    byte[] firstMsoData = dgContainer.getData();

    // Adjust the length of the DgContainer
    int len = IntegerHelper.getInt(firstMsoData[4],
                                   firstMsoData[5],
                                   firstMsoData[6],
                                   firstMsoData[7]);
    IntegerHelper.getFourBytes(len + length, firstMsoData, 4);

    // Adjust the length of the SpgrContainer
    len = IntegerHelper.getInt(firstMsoData[28],
                               firstMsoData[29],
                               firstMsoData[30],
                               firstMsoData[31]);
    IntegerHelper.getFourBytes(len + length, firstMsoData, 28);

    // Now write out each MsoDrawing record

    // First MsoRecord
    // test hack for form objects, to remove the ClientTextBox record
    // from the end of the SpContainer
    if (numImages > 0 &&
        ((DrawingGroupObject) drawings.get(0)).isFormObject())
    {
      byte[] msodata2 = new byte[firstMsoData.length - 8];
      System.arraycopy(firstMsoData, 0, msodata2, 0, msodata2.length);
      firstMsoData = msodata2;
    }

    MsoDrawingRecord msoDrawingRecord = new MsoDrawingRecord(firstMsoData);
    outputFile.write(msoDrawingRecord);

    if (numImages > 0)
    {
      DrawingGroupObject firstDrawing = (DrawingGroupObject) drawings.get(0);
      firstDrawing.writeAdditionalRecords(outputFile);
    }
    else
    {
      // first image is a chart
      Chart chart = charts[0];
      ObjRecord objRecord = chart.getObjRecord();
      outputFile.write(objRecord);
      outputFile.write(chart);
    }

    // Now do all the others
    for (int i = 1; i < spContainerData.length; i++)
    {
      byte[] bytes = (byte[]) spContainerData[i];

      // test hack for form objects, to remove the ClientTextBox record
      // from the end of the SpContainer
      if (i < numImages &&
          ((DrawingGroupObject) drawings.get(i)).isFormObject())
      {
        byte[] bytes2 = new byte[bytes.length - 8];
        System.arraycopy(bytes, 0, bytes2, 0, bytes2.length);
        bytes = bytes2;
      }

      msoDrawingRecord = new MsoDrawingRecord(bytes);
      outputFile.write(msoDrawingRecord);

      if (i < numImages)
      {
        // Write anything else the object needs
        DrawingGroupObject d = (DrawingGroupObject) drawings.get(i);
        d.writeAdditionalRecords(outputFile);
      }
      else
      {
        Chart chart = charts[i - numImages];
        ObjRecord objRecord = chart.getObjRecord();
        outputFile.write(objRecord);
        outputFile.write(chart);
      }
    }

    // Write any tail records that need to be written
    for (Iterator i = drawings.iterator(); i.hasNext();)
    {
      DrawingGroupObject dgo2 = (DrawingGroupObject) i.next();
      dgo2.writeTailRecords(outputFile);
    }
  }

  /**
   * Writes out the drawings and the charts if nothing has been modified
   *
   * @param outputFile the output file
   * @exception IOException
   */
  private void writeUnmodified(File outputFile) throws IOException
  {
    if (charts.length == 0 && drawings.size() == 0)
    {
      // No drawings or charts
      return;
    }
    else if (charts.length == 0 && drawings.size() != 0)
    {
      // If there are no charts, then write out the drawings and return
      for (Iterator i = drawings.iterator(); i.hasNext();)
      {
        DrawingGroupObject d = (DrawingGroupObject) i.next();
        outputFile.write(d.getMsoDrawingRecord());
        d.writeAdditionalRecords(outputFile);
      }

      for (Iterator i = drawings.iterator(); i.hasNext();)
      {
        DrawingGroupObject d = (DrawingGroupObject) i.next();
        d.writeTailRecords(outputFile);
      }
      return;
    }
    else if (drawings.size() == 0 && charts.length != 0)
    {
      // If there are no drawings, then write out the charts and return
      Chart curChart = null;
      for (int i = 0; i < charts.length; i++)
      {
        curChart = charts[i];
        if (curChart.getMsoDrawingRecord() != null)
        {
          outputFile.write(curChart.getMsoDrawingRecord());
        }

        if (curChart.getObjRecord() != null)
        {
          outputFile.write(curChart.getObjRecord());
        }

        outputFile.write(curChart);
      }

      return;
    }

    // There are both charts and drawings - the output
    // drawing group records will need
    // to be re-jigged in order to write the drawings out first, then the
    // charts
    int numDrawings = drawings.size();
    int length = 0;
    EscherContainer[] spContainers =
      new EscherContainer[numDrawings + charts.length];
    boolean[] isFormObject = new boolean[numDrawings + charts.length];

    for (int i = 0; i < numDrawings; i++)
    {
      DrawingGroupObject d = (DrawingGroupObject) drawings.get(i);
      spContainers[i] = d.getSpContainer();

      if (i > 0)
      {
        length += spContainers[i].getLength();
      }

      if (d.isFormObject())
      {
        isFormObject[i] = true;
      }
    }

    for (int i = 0; i < charts.length; i++)
    {
      spContainers[i + numDrawings] = charts[i].getSpContainer();
      length += spContainers[i + numDrawings].getLength();
    }

    // Put the generalised stuff around the first item
    DgContainer dgContainer = new DgContainer();
    Dg dg  = new Dg(numDrawings + charts.length);
    dgContainer.add(dg);

    SpgrContainer spgrContainer = new SpgrContainer();

    SpContainer spContainer = new SpContainer();
    Spgr spgr = new Spgr();
    spContainer.add(spgr);
    Sp sp = new Sp(ShapeType.MIN, 1024, 5);
    spContainer.add(sp);
    spgrContainer.add(spContainer);

    spgrContainer.add(spContainers[0]);

    dgContainer.add(spgrContainer);

    byte[] firstMsoData = dgContainer.getData();

    // Adjust the length of the DgContainer
    int len = IntegerHelper.getInt(firstMsoData[4],
                                   firstMsoData[5],
                                   firstMsoData[6],
                                   firstMsoData[7]);
    IntegerHelper.getFourBytes(len + length, firstMsoData, 4);

    // Adjust the length of the SpgrContainer
    len = IntegerHelper.getInt(firstMsoData[28],
                               firstMsoData[29],
                               firstMsoData[30],
                               firstMsoData[31]);
    IntegerHelper.getFourBytes(len + length, firstMsoData, 28);

    // Now write out each MsoDrawing record and object record

    // Hack to remove the last eight bytes (text box escher record)
    // from the container
    if (isFormObject[0] == true)
    {
      byte[] cbytes = new byte[firstMsoData.length - 8];
      System.arraycopy(firstMsoData, 0, cbytes, 0, cbytes.length);
      firstMsoData = cbytes;
    }

    // First MsoRecord
    MsoDrawingRecord msoDrawingRecord = new MsoDrawingRecord(firstMsoData);
    outputFile.write(msoDrawingRecord);

    DrawingGroupObject dgo = (DrawingGroupObject) drawings.get(0);
    dgo.writeAdditionalRecords(outputFile);

    // Now do all the others
    for (int i = 1; i < spContainers.length; i++)
    {
      byte[] bytes = spContainers[i].getBytes();
      byte[] bytes2 = spContainers[i].setHeaderData(bytes);

      // Hack to remove the last eight bytes (text box escher record)
      // from the container
      if (isFormObject[i] == true)
      {
        byte[] cbytes = new byte[bytes2.length - 8];
        System.arraycopy(bytes2, 0, cbytes, 0, cbytes.length);
        bytes2 = cbytes;
      }

      msoDrawingRecord = new MsoDrawingRecord(bytes2);
      outputFile.write(msoDrawingRecord);

      if (i < numDrawings)
      {
        dgo = (DrawingGroupObject) drawings.get(i);
        dgo.writeAdditionalRecords(outputFile);
      }
      else
      {
        Chart chart = charts[i - numDrawings];
        ObjRecord objRecord = chart.getObjRecord();
        outputFile.write(objRecord);
        outputFile.write(chart);
      }
    }

    // Write any tail records that need to be written
    for (Iterator i = drawings.iterator(); i.hasNext();)
    {
      DrawingGroupObject dgo2 = (DrawingGroupObject) i.next();
      dgo2.writeTailRecords(outputFile);
    }
  }

  /**
   * Sets the charts on the sheet
   *
   * @param ch the charts
   */
  public void setCharts(Chart[] ch)
  {
    charts = ch;
  }

  /**
   * Accessor for the charts on the sheet
   *
   * @return the charts
   */
  public Chart[] getCharts()
  {
    return charts;
  }
}
