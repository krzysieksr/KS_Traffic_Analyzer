import org.opencv.core.*;
import org.opencv.imgproc.Imgproc;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Krzysiek on 2016-07-26.
 */
public class CountVehicles {
    private Mat image;
    public List<MatOfPoint> goodContours = new ArrayList<MatOfPoint>();
    private int areaThreshold;
    private int vehicleSizeThreshold;
    private Point lineCount1;
    private Point lineCount2;
    private Point lineSpeed1;
    private Point lineSpeed2;
    CheckCrossLine checkRectLine;
    CheckCrossLine checkSpeedLine;
    boolean countingFlag = false;
    boolean speedFlag = false;
    boolean crossingLine;
    boolean crossingSpeedLine;
    MatOfPoint contourVehicle;

    public CountVehicles(int areaThreshold, int vehicleSizeThreshold, Point lineCount1, Point lineCount2, Point lineSpeed1, Point lineSpeed2, boolean crossingLine, boolean crossingSpeedLine) {
        this.areaThreshold = areaThreshold;
        this.vehicleSizeThreshold = vehicleSizeThreshold;
        this.lineCount1 = lineCount1;
        this.lineCount2 = lineCount2;
        this.lineSpeed1 = lineSpeed1;
        this.lineSpeed2 = lineSpeed2;
        this.crossingLine = crossingLine;
        this.crossingSpeedLine = crossingSpeedLine;
        this.checkRectLine = new CheckCrossLine(lineCount1, lineCount2);
        this.checkSpeedLine = new CheckCrossLine(lineSpeed1, lineSpeed2);
    }

    public Mat findAndDrawContours(Mat image, Mat binary) {
        List<MatOfPoint> contours = new ArrayList<MatOfPoint>();
        this.image = image;
        Imgproc.findContours(binary, contours, new Mat(), Imgproc.CHAIN_APPROX_NONE, Imgproc.CHAIN_APPROX_SIMPLE);
        Imgproc.line(image, lineCount1, lineCount2, new Scalar(0, 0, 255), 1);
        Imgproc.line(image, lineSpeed1, lineSpeed2, new Scalar(0, 255, 0), 1);

        for (int i = 0; i < contours.size(); i++) {
            MatOfPoint currentContour = contours.get(i);
            double currentArea = Imgproc.contourArea(currentContour);

            if (currentArea > areaThreshold) {
                goodContours.add(contours.get(i));
                drawBoundingBox(currentContour);
            }
        }

        return image;
    }

    public boolean isVehicleToAdd() {
        for (int i = 0; i < goodContours.size(); i++) {
            Rect rectangle = Imgproc.boundingRect(goodContours.get(i));
            if (checkRectLine.rectContainLine(rectangle)) {
                contourVehicle = getGoodContours().get(i);
                countingFlag = true;
                break;
            }
        }
        if (countingFlag == true) {
            if (crossingLine == false) {
                crossingLine = true;
                return true;
            } else {

                return false;
            }
        } else {
            crossingLine = false;
            return false;
        }
    }

    public String classifier() {
        double currentArea = Imgproc.contourArea(contourVehicle);
        if (currentArea <= (double) vehicleSizeThreshold)
            return "Car";
        else if (currentArea <= 1.9 * (double) vehicleSizeThreshold)
            return "Van";
        else return "Lorry";
    }

    private void drawBoundingBox(MatOfPoint currentContour) {
        Rect rectangle = Imgproc.boundingRect(currentContour);
        Imgproc.rectangle(image, rectangle.tl(), rectangle.br(), new Scalar(255, 0, 0), 1);

    }

    public boolean isToSpeedMeasure() {
        for (int i = 0; i < goodContours.size(); i++) {
            Rect rectangle = Imgproc.boundingRect(goodContours.get(i));
            if (checkSpeedLine.rectContainLine(rectangle)) {
                speedFlag = true;
                break;
            }
        }
        if (speedFlag == true) {
            if (crossingSpeedLine == false) {
                crossingSpeedLine = true;
                return true;
            } else {
                return false;
            }
        } else {
            crossingSpeedLine = false;
            return false;
        }
    }

    public boolean isCrossingSpeedLine() {
        return crossingSpeedLine;
    }

    public boolean isCrossingLine() {
        return crossingLine;
    }

    public List<MatOfPoint> getGoodContours() {
        return goodContours;
    }

}
