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


package jxl.format;

/**
 * A structure which contains the RGB values for a particular colour
 */
public final class RGB
{
  /**
   * The red component of this colour
   */
  private  int red;

  /**
   * The green component of this colour
   */
  private int green;

  /**
   * The blue component of this colour
   */
  private int blue;

  /**
   * Constructor
   *
   * @param r the red component
   * @param g the green component
   * @param b the blue component
   */
  public RGB(int r, int g, int b)
  {
    red = r;
    green = g;
    blue = b;
  }

  /**
   * Accessor for the red component
   *
   * @return the red component of the colour, between 0 and 255
   */
  public int getRed()
  {
    return red;
  }

  /**
   * Accessor for the green component
   *
   * @return the green component of the colour, between 0 and 255
   */
  public int getGreen()
  {
    return green;
  }

  /**
   * Accessor for the blue component
   *
   * @return the blue component of the colour, between 0 and 255
   */
  public int getBlue()
  {
    return blue;
  }
}

