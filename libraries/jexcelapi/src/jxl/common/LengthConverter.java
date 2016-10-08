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

package jxl.common;

public class LengthConverter
{
  private static double[][] factors = 
    new double[LengthUnit.getCount()][LengthUnit.getCount()];

  static
  {
    // The identity factors
    factors[LengthUnit.POINTS.getIndex()][LengthUnit.POINTS.getIndex()] = 1;
    factors[LengthUnit.METRES.getIndex()][LengthUnit.METRES.getIndex()] = 1;
    factors[LengthUnit.CENTIMETRES.getIndex()][LengthUnit.CENTIMETRES.getIndex()] = 1;
    factors[LengthUnit.INCHES.getIndex()][LengthUnit.INCHES.getIndex()] = 1;

    // The points conversion factors
    factors[LengthUnit.POINTS.getIndex()][LengthUnit.METRES.getIndex()] = 0.00035277777778;
    factors[LengthUnit.POINTS.getIndex()][LengthUnit.CENTIMETRES.getIndex()] = 0.035277777778;
    factors[LengthUnit.POINTS.getIndex()][LengthUnit.INCHES.getIndex()] = 0.013888888889;

    // The metres conversion factors
    factors[LengthUnit.METRES.getIndex()][LengthUnit.POINTS.getIndex()] = 2877.84;
    factors[LengthUnit.METRES.getIndex()][LengthUnit.CENTIMETRES.getIndex()] = 100;
    factors[LengthUnit.METRES.getIndex()][LengthUnit.INCHES.getIndex()] = 39.37;

    // The centimetres conversion factors
    factors[LengthUnit.CENTIMETRES.getIndex()][LengthUnit.POINTS.getIndex()] = 28.34643;
    factors[LengthUnit.CENTIMETRES.getIndex()][LengthUnit.METRES.getIndex()] = 0.01;
    factors[LengthUnit.CENTIMETRES.getIndex()][LengthUnit.INCHES.getIndex()] = 0.3937;

    // The inches conversion factors
    factors[LengthUnit.INCHES.getIndex()][LengthUnit.POINTS.getIndex()] = 72;
    factors[LengthUnit.INCHES.getIndex()][LengthUnit.METRES.getIndex()] = 0.0254;
    factors[LengthUnit.INCHES.getIndex()][LengthUnit.CENTIMETRES.getIndex()] = 2.54;
  }

  public static double getConversionFactor(LengthUnit from, LengthUnit to)
  {
    return factors[from.getIndex()][to.getIndex()];
  }
}
