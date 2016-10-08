/*********************************************************************
*
*      Copyright (C) 2009 Andrew Khan
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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.WorkbookSettings;
import jxl.biff.ContinueRecord;
import jxl.write.biff.File;

/**
 * Contains the various biff records used to copy a CheckBox (from the
 * Form toolbox) between workbook
 */
public class CheckBox implements DrawingGroupObject
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CheckBox.class);

  /**
   * The spContainer that was read in
   */
  private EscherContainer readSpContainer;

  /**
   * The spContainer that was generated
   */
  private EscherContainer spContainer;

  /**
   * The MsoDrawingRecord associated with the drawing
   */
  private MsoDrawingRecord msoDrawingRecord;

  /**
   * The ObjRecord associated with the drawing
   */
  private ObjRecord objRecord;

  /**
   * Initialized flag
   */
  private boolean initialized = false;

  /**
   * The object id, assigned by the drawing group
   */
  private int objectId;

  /**
   * The blip id
   */
  private int blipId;

  /**
   * The shape id
   */
  private int shapeId;

  /**
   * The column
   */
  private int column;

  /**
   * The row position of the image
   */
  private int row;

  /**
   * The width of the image in cells
   */
  private double width;

  /**
   * The height of the image in cells
   */
  private double height;

  /**
   * The number of places this drawing is referenced
   */
  private int referenceCount;

  /**
   * The top level escher container
   */
  private EscherContainer escherData;

  /**
   * Where this image came from (read, written or a copy)
   */
  private Origin origin;

  /**
   * The drawing group for all the images
   */
  private DrawingGroup drawingGroup;

  /**
   * The drawing data
   */
  private DrawingData drawingData;

  /**
   * The type of this drawing object
   */
  private ShapeType type;

  /**
   * The drawing position on the sheet
   */
  private int drawingNumber;

  /**
   * An mso drawing record, which sometimes appears
   */
  private MsoDrawingRecord mso;

  /**
   * The text object record
   */
  private TextObjectRecord txo;

  /**
   * Text data from the first continue record
   */
  private ContinueRecord text;

  /**
   * Formatting data from the second continue record
   */
  private ContinueRecord formatting;

  /**
   * The workbook settings
   */
  private WorkbookSettings workbookSettings;

  /**
   * Constructor used when reading images
   *
   * @param mso the drawing record
   * @param obj the object record
   * @param dd the drawing data for all drawings on this sheet
   * @param dg the drawing group
   * @param ws the workbook settings
   */
  public CheckBox(MsoDrawingRecord mso, ObjRecord obj, DrawingData dd,
                  DrawingGroup dg, WorkbookSettings ws)
  {
    drawingGroup = dg;
    msoDrawingRecord = mso;
    drawingData = dd;
    objRecord = obj;
    initialized = false;
    workbookSettings = ws;
    origin = Origin.READ;
    drawingData.addData(msoDrawingRecord.getData());
    drawingNumber = drawingData.getNumDrawings() - 1;
    drawingGroup.addDrawing(this);

    Assert.verify(mso != null && obj != null);

    initialize();
  }

  /**
   * Copy constructor used to copy drawings from read to write
   *
   * @param dgo the drawing group object
   * @param dg the drawing group
   * @param ws the workbook settings
   */
  public CheckBox(DrawingGroupObject dgo,
                  DrawingGroup dg,
                  WorkbookSettings ws)
  {
    CheckBox d = (CheckBox) dgo;
    Assert.verify(d.origin == Origin.READ);
    msoDrawingRecord = d.msoDrawingRecord;
    objRecord = d.objRecord;
    initialized = false;
    origin = Origin.READ;
    drawingData = d.drawingData;
    drawingGroup = dg;
    drawingNumber = d.drawingNumber;
    drawingGroup.addDrawing(this);
    mso = d.mso;
    txo = d.txo;
    text = d.text;
    formatting = d.formatting;
    workbookSettings = ws;
  }

  /**
   * Constructor invoked when writing images
   */
  public CheckBox()
  {
    initialized = true;
    origin = Origin.WRITE;
    referenceCount = 1;
    type = ShapeType.HOST_CONTROL;
  }

  /**
   * Initializes the member variables from the Escher stream data
   */
  private void initialize()
  {
    readSpContainer = drawingData.getSpContainer(drawingNumber);
    Assert.verify(readSpContainer != null);

    EscherRecord[] children = readSpContainer.getChildren();

    Sp sp = (Sp) readSpContainer.getChildren()[0];
    objectId = objRecord.getObjectId();
    shapeId = sp.getShapeId();
    type = ShapeType.getType(sp.getShapeType());

    if (type == ShapeType.UNKNOWN)
    {
      logger.warn("Unknown shape type");
    }

    ClientAnchor clientAnchor = null;
    for (int i = 0; i < children.length && clientAnchor == null; i++)
    {
      if (children[i].getType() == EscherRecordType.CLIENT_ANCHOR)
      {
        clientAnchor = (ClientAnchor) children[i];
      }
    }

    if (clientAnchor == null)
    {
      logger.warn("Client anchor not found");
    }
    else
    {
      column = (int) clientAnchor.getX1();
      row = (int) clientAnchor.getY1();
    }

    initialized = true;
  }


  /**
   * Sets the object id.  Invoked by the drawing group when the object is
   * added to id
   *
   * @param objid the object id
   * @param bip the blip id
   * @param sid the shape id
   */
  public final void setObjectId(int objid, int bip, int sid)
  {
    objectId = objid;
    blipId = bip;
    shapeId = sid;

    if (origin == Origin.READ)
    {
      origin = Origin.READ_WRITE;
    }
  }

  /**
   * Accessor for the object id
   *
   * @return the object id
   */
  public final int getObjectId()
  {
    if (!initialized)
    {
      initialize();
    }

    return objectId;
  }

  /**
   * Accessor for the shape id
   *
   * @return the object id
   */
  public final int getShapeId()
  {
    if (!initialized)
    {
      initialize();
    }

    return shapeId;
  }

  /**
   * Accessor for the blip id
   *
   * @return the blip id
   */
  public final int getBlipId()
  {
    if (!initialized)
    {
      initialize();
    }

    return blipId;
  }

  /**
   * Gets the drawing record which was read in
   *
   * @return the drawing record
   */
  public MsoDrawingRecord  getMsoDrawingRecord()
  {
    return msoDrawingRecord;
  }

  /**
   * Creates the main Sp container for the drawing
   *
   * @return the SP container
   */
  public EscherContainer getSpContainer()
  {
    if (!initialized)
    {
      initialize();
    }

    if (origin == Origin.READ)
    {
      return getReadSpContainer();
    }

    SpContainer spc = new SpContainer();
    Sp sp = new Sp(type, shapeId, 2560);
    spc.add(sp);
    Opt opt = new Opt();
    opt.addProperty(127, false, false, 17039620);
    opt.addProperty(191, false, false, 524296);
    opt.addProperty(511, false, false, 524288);
    opt.addProperty(959, false, false, 131072);
    //    opt.addProperty(260, true, false, blipId);
    //    opt.addProperty(261, false, false, 36);
    spc.add(opt);

    ClientAnchor clientAnchor = new ClientAnchor(column,
                                                 row,
                                                 column + 1,
                                                 row + 1, 
                                                 0x1);
    spc.add(clientAnchor);
    ClientData clientData = new ClientData();
    spc.add(clientData);

    return spc;
  }

  /**
   * Sets the drawing group for this drawing.  Called by the drawing group
   * when this drawing is added to it
   *
   * @param dg the drawing group
   */
  public void setDrawingGroup(DrawingGroup dg)
  {
    drawingGroup = dg;
  }

  /**
   * Accessor for the drawing group
   *
   * @return the drawing group
   */
  public DrawingGroup getDrawingGroup()
  {
    return drawingGroup;
  }

  /**
   * Gets the origin of this drawing
   *
   * @return where this drawing came from
   */
  public Origin getOrigin()
  {
    return origin;
  }

  /**
   * Accessor for the reference count on this drawing
   *
   * @return the reference count
   */
  public int getReferenceCount()
  {
    return referenceCount;
  }

  /**
   * Sets the new reference count on the drawing
   *
   * @param r the new reference count
   */
  public void setReferenceCount(int r)
  {
    referenceCount = r;
  }

  /**
   * Accessor for the column of this drawing
   *
   * @return the column
   */
  public double getX()
  {
    if (!initialized)
    {
      initialize();
    }
    return column;
  }

  /**
   * Sets the column position of this drawing.  Used when inserting/removing
   * columns from the spreadsheet
   *
   * @param x the column
   */
  public void setX(double x)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    column = (int) x;
  }

  /**
   * Accessor for the row of this drawing
   *
   * @return the row
   */
  public double getY()
  {
    if (!initialized)
    {
      initialize();
    }

    return row;
  }

  /**
   * Accessor for the row of the drawing
   *
   * @param y the row
   */
  public void setY(double y)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    row = (int) y;
  }


  /**
   * Accessor for the width of this drawing
   *
   * @return the number of columns spanned by this image
   */
  public double getWidth()
  {
    if (!initialized)
    {
      initialize();
    }

    return width;
  }

  /**
   * Accessor for the width
   *
   * @param w the number of columns to span
   */
  public void setWidth(double w)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    width = w;
  }

  /**
   * Accessor for the height of this drawing
   *
   * @return the number of rows spanned by this image
   */
  public double getHeight()
  {
    if (!initialized)
    {
      initialize();
    }

    return height;
  }

  /**
   * Accessor for the height of this drawing
   *
   * @param h the number of rows spanned by this image
   */
  public void setHeight(double h)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    height = h;
  }


  /**
   * Gets the SpContainer that was read in
   *
   * @return the read sp container
   */
  private EscherContainer getReadSpContainer()
  {
    if (!initialized)
    {
      initialize();
    }

    return readSpContainer;
  }

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  public byte[] getImageData()
  {
    Assert.verify(origin == Origin.READ || origin == Origin.READ_WRITE);

    if (!initialized)
    {
      initialize();
    }

    return drawingGroup.getImageData(blipId);
  }

  /**
   * Accessor for the type
   *
   * @return the type
   */
  public ShapeType getType()
  {
    return type;
  }

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  public byte[] getImageBytes()
  {
    Assert.verify(false);
    return null;
  }

  /**
   * Accessor for the image file path.  Normally this is the absolute path
   * of a file on the directory system, but if this drawing was constructed
   * using an byte[] then the blip id is returned
   *
   * @return the image file path, or the blip id
   */
  public String getImageFilePath()
  {
    Assert.verify(false);
    return null;
  }

  /**
   * Writes out the additional records for a combo box
   *
   * @param outputFile the output file
   * @exception IOException
   */
  public void writeAdditionalRecords(File outputFile) throws IOException
  {
    if (origin == Origin.READ)
    {
      outputFile.write(objRecord);

      if (mso != null)
      {
        outputFile.write(mso);
      }
      outputFile.write(txo);
      outputFile.write(text);
      if (formatting != null)
      {
        outputFile.write(formatting);
      }
      return;
    }

    // Create the obj record
    ObjRecord objrec = new ObjRecord(objectId,
                                     ObjRecord.CHECKBOX);

    outputFile.write(objrec);

    logger.warn("Writing of additional records for checkboxes not " +
                "implemented");
  }

  /**
   * Writes any records that need to be written after all the drawing group
   * objects have been written
   * Writes out all the note records
   *
   * @param outputFile the output file
   */
  public void writeTailRecords(File outputFile)
  {
  }

  /**
   * Accessor for the row
   *
   * @return  the row
   */
  public int getRow()
  {
    return 0;
  }

  /**
   * Accessor for the column
   *
   * @return the column
   */
  public int getColumn()
  {
    return 0;
  }

  /**
   * Hashing algorithm
   *
   * @return  the hash code
   */
  public int hashCode()
  {
    return getClass().getName().hashCode();
  }

  /**
   * Accessor for the first drawing on the sheet.  This is used when
   * copying unmodified sheets to indicate that this drawing contains
   * the first time Escher gubbins
   *
   * @return TRUE if this MSORecord is the first drawing on the sheet
   */
  public boolean isFirst()
  {
    return msoDrawingRecord.isFirst();
  }

  /**
   * Queries whether this object is a form object.  Form objects have their
   * drawings records spread over TXO and CONTINUE records and
   * require special handling
   *
   * @return TRUE if this is a form object, FALSE otherwise
   */
  public boolean isFormObject()
  {
    return false;
  }

  /**
   * Sets the text object
   *
   * @param t the text object record
   */
  public void setTextObject(TextObjectRecord t)
  {
    txo = t;
  }

  /**
   * Sets the text data
   *
   * @param t continuation record
   */
  public void setText(ContinueRecord t)
  {
    text = t;
  }

  /**
   * Sets the formatting
   *
   * @param t continue record
   */
  public void setFormatting(ContinueRecord t)
  {
    formatting = t;
  }

  /**
   * The drawing record
   *
   * @param d the drawing record
   */
  public void addMso(MsoDrawingRecord d)
  {
    mso = d;
    drawingData.addRawData(mso.getData());
  }

}



