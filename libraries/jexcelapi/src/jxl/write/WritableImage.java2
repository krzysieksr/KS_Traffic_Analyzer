/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

package jxl.write;

import java.io.File;

import jxl.biff.drawing.Drawing;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.biff.drawing.DrawingGroup;

/**
 * Allows an image to be created, or an existing image to be manipulated
 * Note that co-ordinates and dimensions are given in cells, so that if for
 * example the width or height of a cell which the image spans is altered,
 * the image will have a correspondign distortion
 */
public class WritableImage extends Drawing
{
  /**
   * Constructor
   *
   * @param x the column number at which to position the image
   * @param y the row number at which to position the image
   * @param width the number of columns cells which the image spans
   * @param height the number of rows which the image spans
   * @param image the source image file
   */
  public WritableImage(double x, double y,
                       double width, double height,
                       File image)
  {
    super(x, y, width, height, image);
  }

  /**
   * Constructor
   *
   * @param x the column number at which to position the image
   * @param y the row number at which to position the image
   * @param width the number of columns cells which the image spans
   * @param height the number of rows which the image spans
   * @param image the image data
   */
  public WritableImage(double x, double y,
                       double width, double height,
                       byte[] imageData)
  {
    super(x, y, width, height, imageData);
  }

  /**
   * Constructor, used when copying sheets
   *
   * @param d the image to copy
   */
  public WritableImage(DrawingGroupObject d, DrawingGroup dg)
  {
    super(d, dg);
  }

  /**
   * Accessor for the image position
   *
   * @return the column number at which the image is positioned
   */
  public double getColumn()
  {
    return super.getX();
  }

  /**
   * Accessor for the image position
   *
   * @param c the column number at which the image should be positioned
   */
  public void setColumn(double c)
  {
    super.setX(c);
  }

  /**
   * Accessor for the image position
   *
   * @return the row number at which the image is positions
   */
  public double getRow()
  {
    return super.getY();
  }

  /**
   * Accessor for the image position
   *
   * @param c the row number at which the image should be positioned
   */
  public void setRow(double c)
  {
    super.setY(c);
  }

  /**
   * Accessor for the image dimensions
   *
   * @return  the number of columns this image spans
   */
  public double getWidth()
  {
    return super.getWidth();
  }

  /**
   * Accessor for the image dimensions
   * Note that the actual size of the rendered image will depend on the
   * width of the columns it spans
   *
   * @param c the number of columns which this image spans
   */
  public void setWidth(double c)
  {
    super.setWidth(c);
  }

  /**
   * Accessor for the image dimensions
   *
   * @return the number of rows which this image spans
   */
  public double getHeight()
  {
    return super.getHeight();
  }

  /**
   * Accessor for the image dimensions
   * Note that the actual size of the rendered image will depend on the
   * height of the rows it spans
   *
   * @param c the number of rows which this image should span
   */
  public void setHeight(double c)
  {
    super.setHeight(c);
  }

  /**
   * Accessor for the image file
   *
   * @return the file which the image references
   */
  public File getImageFile()
  {
    return super.getImageFile();
  }

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  public byte[] getImageData()
  {
    return super.getImageData();
  }
}
